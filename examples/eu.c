// /Users/tamc/Documents/github/excel_to_code/examples/eu.xlsx approximately translated into C
// First we have c versions of all the excel functions that we know
#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include <math.h>

// FIXME: Extract a header file

// I predefine an array of ExcelValues to store calculations
// Probably bad practice. At the very least, I should make it
// link to the cell reference in some way.
#define MAX_EXCEL_VALUE_HEAP_SIZE 1000000

#define true 1
#define false 0

// These are the various types of excel cell, plus ExcelRange which allows the passing of arrays of cells
typedef enum {ExcelEmpty, ExcelNumber, ExcelString, ExcelBoolean, ExcelError, ExcelRange} ExcelType;

struct excel_value {
	ExcelType type;
	
	double number; // Used for numbers and for error types
	char *string; // Used for strings
	
	// The following three are used for ranges
	void *array;
	int rows;
	int columns;
};

typedef struct excel_value ExcelValue;


// These are used in the SUMIF and SUMIFS criteria (e.g., when passed a string like "<20")
typedef enum {LessThan, LessThanOrEqual, Equal, NotEqual, MoreThanOrEqual, MoreThan} ExcelComparisonType;

struct excel_comparison {
	ExcelComparisonType type;
	ExcelValue comparator;
};

typedef struct excel_comparison ExcelComparison;

// Headers
static ExcelValue more_than(ExcelValue a_v, ExcelValue b_v);
static ExcelValue more_than_or_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue not_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue less_than(ExcelValue a_v, ExcelValue b_v);
static ExcelValue less_than_or_equal(ExcelValue a_v, ExcelValue b_v);
static ExcelValue find_2(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v);
static ExcelValue find(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v, ExcelValue position_to_start_at_v);
static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error);
static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v);
static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue row_number_v);
static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v);
static ExcelValue left_1(ExcelValue string_v);
static ExcelValue max(int number_of_arguments, ExcelValue *arguments);
static ExcelValue min(int number_of_arguments, ExcelValue *arguments);
static ExcelValue mod(ExcelValue a_v, ExcelValue b_v);
static ExcelValue negative(ExcelValue a_v);
static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v);
static ExcelValue power(ExcelValue a_v, ExcelValue b_v);
static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments);
static ExcelValue subtotal(ExcelValue type, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumif(ExcelValue check_range_v, ExcelValue criteria_v, ExcelValue sum_range_v );
static ExcelValue sumif_2(ExcelValue check_range_v, ExcelValue criteria_v);
static ExcelValue sumproduct(int number_of_arguments, ExcelValue *arguments);
static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v);
static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v);

// My little heap
ExcelValue cells[MAX_EXCEL_VALUE_HEAP_SIZE];
int cell_counter = 0;

#define HEAPCHECK if(cell_counter >= MAX_EXCEL_VALUE_HEAP_SIZE) { printf("Heap exceeded"); exit(-1); }

// The object initializers
static ExcelValue new_excel_number(double number) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelNumber;
	new_cell.number = number;
	return new_cell;
};

static ExcelValue new_excel_string(char *string) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelString;
	new_cell.string = string;
	return new_cell;
};

static ExcelValue new_excel_range(void *array, int rows, int columns) {
	cell_counter++;
	HEAPCHECK
	ExcelValue new_cell = cells[cell_counter];
	new_cell.type = ExcelRange;
	new_cell.array = array;
	new_cell.rows = rows;
	new_cell.columns = columns;
	return new_cell;
};

static void * new_excel_value_array(int size) {
	ExcelValue *pointer = malloc(sizeof(ExcelValue)*size);
	if(pointer == 0) {
		printf("Out of memory\n");
		exit(-1);
	}
	return pointer;
};

// Constants
const ExcelValue BLANK = {.type = ExcelEmpty, .number = 0};

const ExcelValue ZERO = {.type = ExcelNumber, .number = 0};
const ExcelValue ONE = {.type = ExcelNumber, .number = 1};
const ExcelValue TWO = {.type = ExcelNumber, .number = 2};
const ExcelValue THREE = {.type = ExcelNumber, .number = 3};
const ExcelValue FOUR = {.type = ExcelNumber, .number = 4};
const ExcelValue FIVE = {.type = ExcelNumber, .number = 5};
const ExcelValue SIX = {.type = ExcelNumber, .number = 6};
const ExcelValue SEVEN = {.type = ExcelNumber, .number = 7};
const ExcelValue EIGHT = {.type = ExcelNumber, .number = 8};
const ExcelValue NINE = {.type = ExcelNumber, .number = 9};
const ExcelValue TEN = {.type = ExcelNumber, .number = 10};


// Booleans
const ExcelValue TRUE = {.type = ExcelBoolean, .number = true };
const ExcelValue FALSE = {.type = ExcelBoolean, .number = false };

// Errors
const ExcelValue VALUE = {.type = ExcelError, .number = 0};
const ExcelValue NAME = {.type = ExcelError, .number = 1};
const ExcelValue DIV0 = {.type = ExcelError, .number = 2};
const ExcelValue REF = {.type = ExcelError, .number = 3};
const ExcelValue NA = {.type = ExcelError, .number = 4};

// This is the error flag
static int conversion_error = 0;

// Helpful for debugging
static void inspect_excel_value(ExcelValue v) {
	ExcelValue *array;
	int i, j, k;
	switch (v.type) {
  	  case ExcelNumber:
		  printf("Number: %f\n",v.number);
		  break;
	  case ExcelBoolean:
		  if(v.number == true) {
			  printf("True\n");
		  } else if(v.number == false) {
			  printf("False\n");
		  } else {
			  printf("Boolean with undefined state %f\n",v.number);
		  }
		  break;
	  case ExcelEmpty:
	  	if(v.number == 0) {
	  		printf("Empty\n");
		} else {
			printf("Empty with unexpected state %f\n",v.number);	
		}
		break;
	  case ExcelRange:
		 printf("Range rows: %d, columns: %d\n",v.rows,v.columns);
		 array = v.array;
		 for(i = 0; i < v.rows; i++) {
			 printf("Row %d:\n",i+1);
			 for(j = 0; j < v.columns; j++ ) {
				 printf("%d ",j+1);
				 k = (i * v.columns) + j;
				 inspect_excel_value(array[k]);
			 }
		 }
		 break;
	  case ExcelString:
		 printf("String: '%s'\n",v.string);
		 break;
	  case ExcelError:
		 printf("Error number %f ",v.number);
		 switch( (int)v.number) {
			 case 0: printf("VALUE\n"); break;
			 case 1: printf("NAME\n"); break;
			 case 2: printf("DIV0\n"); break;
			 case 3: printf("REF\n"); break;
			 case 4: printf("NA\n"); break;
		 }
		 break;
    default:
      printf("Type %d not recognised",v.type);
	 };
}

// Extracts numbers from ExcelValues
// Excel treats empty cells as zero
static double number_from(ExcelValue v) {
	char *s;
	char *p;
	double n;
	ExcelValue *array;
	switch (v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  	return v.number;
	  case ExcelEmpty: 
	  	return 0;
	  case ExcelRange: 
		 array = v.array;
	     return number_from(array[0]);
	  case ExcelString:
 	 	s = v.string;
		if (s == NULL || *s == '\0' || isspace(*s)) {
			return 0;
		}	        
		n = strtod (s, &p);
		if(*p == '\0') {
			return n;
		}
		conversion_error = 1;
		return 0;
	  case ExcelError:
	  	return 0;
  }
  return 0;
}

#define NUMBER(value_name, name) double name; if(value_name.type == ExcelError) { return value_name; }; name = number_from(value_name);
#define CHECK_FOR_CONVERSION_ERROR 	if(conversion_error) { conversion_error = 0; return VALUE; };
#define CHECK_FOR_PASSED_ERROR(name) 	if(name.type == ExcelError) return name;
	
static ExcelValue excel_abs(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)	
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	
	if(a >= 0.0 ) {
		return a_v;
	} else {
		return new_excel_number(-a);
	}
}

static ExcelValue add(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a + b);
}

static ExcelValue excel_and(int array_size, ExcelValue *array) {
	int i;
	ExcelValue current_excel_value, array_result;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber: 
		  case ExcelBoolean: 
			  if(current_excel_value.number == false) return FALSE;
			  break;
		  case ExcelRange: 
		  	array_result = excel_and( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(array_result.type == ExcelError) return array_result;
			if(array_result.type == ExcelBoolean && array_result.number == false) return FALSE;
			break;
		  case ExcelString:
		  case ExcelEmpty:
			 break;
		  case ExcelError:
			 return current_excel_value;
			 break;
		 }
	 }
	 return TRUE;
}

struct average_result {
	double sum;
	double count;
	int has_error;
	ExcelValue error;
};
	
static struct average_result calculate_average(int array_size, ExcelValue *array) {
	double sum = 0;
	double count = 0;
	int i;
	ExcelValue current_excel_value;
	struct average_result array_result, r;
		 
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber:
			  sum += current_excel_value.number;
			  count++;
			  break;
		  case ExcelRange: 
		  	array_result = calculate_average( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(array_result.has_error == true) return array_result;
			sum += array_result.sum;
			count += array_result.count;
			break;
		  case ExcelBoolean: 
		  case ExcelString:
		  case ExcelEmpty:
			 break;
		  case ExcelError:
			 r.has_error = true;
			 r.error = current_excel_value;
			 return r;
			 break;
		 }
	}
	r.count = count;
	r.sum = sum;
	r.has_error = false;
	return r;
}

static ExcelValue average(int array_size, ExcelValue *array) {
	struct average_result r = calculate_average(array_size, array);
	if(r.has_error == true) return r.error;
	if(r.count == 0) return DIV0;
	return new_excel_number(r.sum/r.count);
}

static ExcelValue choose(ExcelValue index_v, int array_size, ExcelValue *array) {
	CHECK_FOR_PASSED_ERROR(index_v)

	int index = (int) number_from(index_v);
	CHECK_FOR_CONVERSION_ERROR	
	int i;
	for(i=0;i<array_size;i++) {
		if(array[i].type == ExcelError) return array[i];
	}
	if(index < 1) return VALUE;
	if(index > array_size) return VALUE;
	return array[index-1];
}	

static ExcelValue count(int array_size, ExcelValue *array) {
	int i;
	int n = 0;
	ExcelValue current_excel_value;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber:
		  	n++;
			break;
		  case ExcelRange: 
		  	n += count( current_excel_value.rows * current_excel_value.columns, current_excel_value.array ).number;
			break;
  		  case ExcelBoolean: 			
		  case ExcelString:
		  case ExcelEmpty:
		  case ExcelError:
			 break;
		 }
	 }
	 return new_excel_number(n);
}

static ExcelValue counta(int array_size, ExcelValue *array) {
	int i;
	int n = 0;
	ExcelValue current_excel_value;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
    switch(current_excel_value.type) {
  	  case ExcelNumber:
      case ExcelBoolean:
      case ExcelString:
  	  case ExcelError:
        n++;
        break;
      case ExcelRange: 
	  	  n += counta( current_excel_value.rows * current_excel_value.columns, current_excel_value.array ).number;
        break;
  	  case ExcelEmpty:
  		  break;
    }
	 }
	 return new_excel_number(n);
}

static ExcelValue divide(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return new_excel_number(a / b);
}

static ExcelValue excel_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	if(a_v.type != b_v.type) return FALSE;
	
	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty: 
			if(a_v.number != b_v.number) return FALSE;
			return TRUE;
	  case ExcelString:
	  	if(strcasecmp(a_v.string,b_v.string) != 0 ) return FALSE;
		return TRUE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue not_equal(ExcelValue a_v, ExcelValue b_v) {
	ExcelValue result = excel_equal(a_v, b_v);
	if(result.type == ExcelBoolean) {
		if(result.number == 0) return TRUE;
		return FALSE;
	}
	return result;
}

static ExcelValue excel_if(ExcelValue condition, ExcelValue true_case, ExcelValue false_case ) {
	CHECK_FOR_PASSED_ERROR(condition)
	CHECK_FOR_PASSED_ERROR(true_case)
	CHECK_FOR_PASSED_ERROR(false_case)
	
	switch (condition.type) {
  	  case ExcelBoolean:
  	  	if(condition.number == true) return true_case;
  	  	return false_case;
  	  case ExcelNumber:
		if(condition.number == false) return false_case;
		return true_case;
	  case ExcelEmpty: 
		return false_case;
	  case ExcelString:
	  	return VALUE;
  	  case ExcelError:
		return condition;
  	  case ExcelRange:
  		return VALUE;
  }
  return condition;
}

static ExcelValue excel_if_2(ExcelValue condition, ExcelValue true_case ) {
	return excel_if( condition, true_case, FALSE );
}

static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v) {
	CHECK_FOR_PASSED_ERROR(array_v)
	CHECK_FOR_PASSED_ERROR(row_number_v)
	CHECK_FOR_PASSED_ERROR(column_number_v)
		
	ExcelValue *array;
	int rows;
	int columns;
	
	NUMBER(row_number_v, row_number)
	NUMBER(column_number_v, column_number)
	CHECK_FOR_CONVERSION_ERROR
	
	if(array_v.type == ExcelRange) {
		array = array_v.array;
		rows = array_v.rows;
		columns = array_v.columns;
	} else {
		ExcelValue tmp_array[] = {array_v};
		array = tmp_array;
		rows = 1;
		columns = 1;
	}
	
	if(row_number > rows) return REF;
	if(column_number > columns) return REF;
		
	if(row_number == 0) { // We need the whole column
		if(column_number < 1) return REF;
		ExcelValue *result = (ExcelValue *) new_excel_value_array(rows);
		int result_index = 0;
		ExcelValue r;
		int array_index;
		int i;
		for(i = 0; i < rows; i++) {
			array_index = (i*columns) + column_number - 1;
			r = array[array_index];
			if(r.type == ExcelEmpty) {
				result[result_index] = ZERO;
			} else {
				result[result_index] = r;
			}			
			result_index++;
		}
		return new_excel_range(result,rows,1);
	} else if(column_number == 0 ) { // We need the whole row
		if(row_number < 1) return REF;
		ExcelValue *result = (ExcelValue*) new_excel_value_array(columns);
		ExcelValue r;
		int row_start = ((row_number-1)*columns);
		int row_finish = row_start + columns;
		int result_index = 0;
		int i;
		for(i = row_start; i < row_finish; i++) {
			r = array[i];
			if(r.type == ExcelEmpty) {
				result[result_index] = ZERO;
			} else {
				result[result_index] = r;
			}
			result_index++;
		}
		return new_excel_range(result,1,columns);
	} else { // We need a precise point
		if(row_number < 1) return REF;
		if(column_number < 1) return REF;
		int position = ((row_number - 1) * columns) + column_number - 1;
		ExcelValue result = array[position];
		if(result.type == ExcelEmpty) return ZERO;
		return result;
	}
	
	return FALSE;
};

static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue offset) {
	if(array_v.type == ExcelRange) {
		if(array_v.rows == 1) {
			return excel_index(array_v,ONE,offset);
		} else if (array_v.columns == 1) {
			return excel_index(array_v,offset,ONE);
		} else {
			return REF;
		}
	} else if (offset.type == ExcelNumber && offset.number == 1) {
		return array_v;
	} else {
		return REF;
	}
	return REF;
};


static ExcelValue excel_match(ExcelValue lookup_value, ExcelValue lookup_array, ExcelValue match_type ) {
	CHECK_FOR_PASSED_ERROR(lookup_value)
	CHECK_FOR_PASSED_ERROR(lookup_array)
	CHECK_FOR_PASSED_ERROR(match_type)
		
	// Blanks are treaked as zeros
	if(lookup_value.type == ExcelEmpty) lookup_value = ZERO;

	// Setup the array
	ExcelValue *array;
	int size;
	if(lookup_array.type == ExcelRange) {
		// Check that the range is a row or column rather than an area
		if((lookup_array.rows == 1) || (lookup_array.columns == 1)) {
			array = lookup_array.array;
			size = lookup_array.rows * lookup_array.columns;
		} else {
			// return NA error if covers an area.
			return NA;
		};
	} else {
		// Need to wrap the argument up as an array
		size = 1;
		ExcelValue tmp_array[1] = {lookup_array};
		array = tmp_array;
	}
    
	int type = (int) number_from(match_type);
	CHECK_FOR_CONVERSION_ERROR;
	
	int i;
	ExcelValue x;
	
	switch(type) {
		case 0:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(excel_equal(lookup_value,x).number == true) return new_excel_number(i+1);
			}
			return NA;
			break;
		case 1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(more_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return new_excel_number(i);
				}
			}
			return new_excel_number(size);
			break;
		case -1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(x.type == ExcelEmpty) x = ZERO;
				if(less_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return new_excel_number(i);
				}
			}
			return new_excel_number(size-1);
			break;
	}
	return NA;
}

static ExcelValue excel_match_2(ExcelValue lookup_value, ExcelValue lookup_array ) {
	return excel_match(lookup_value,lookup_array,new_excel_number(0));
}

static ExcelValue find(ExcelValue find_text_v, ExcelValue within_text_v, ExcelValue start_number_v) {
	CHECK_FOR_PASSED_ERROR(find_text_v)
	CHECK_FOR_PASSED_ERROR(within_text_v)
	CHECK_FOR_PASSED_ERROR(start_number_v)

	char *find_text;	
	char *within_text;
	char *within_text_offset;
	char *result;
	int start_number = number_from(start_number_v);
	CHECK_FOR_CONVERSION_ERROR

	// Deal with blanks 
	if(within_text_v.type == ExcelString) {
		within_text = within_text_v.string;
	} else if( within_text_v.type == ExcelEmpty) {
		within_text = "";
	}

	if(find_text_v.type == ExcelString) {
		find_text = find_text_v.string;
	} else if( find_text_v.type == ExcelEmpty) {
		return start_number_v;
	}
	
	// Check length
	if(start_number < 1) return VALUE;
	if(start_number > strlen(within_text)) return VALUE;
	
	// Offset our within_text pointer
	// FIXME: No way this is utf-8 compatible
	within_text_offset = within_text + (start_number - 1);
	result = strstr(within_text_offset,find_text);
	if(result) {
		return new_excel_number(result - within_text + 1);
	}
	return VALUE;
}

static ExcelValue find_2(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v) {
	return find(string_to_look_for_v, string_to_look_in_v, ONE);
};

static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v) {
	CHECK_FOR_PASSED_ERROR(string_v)
	CHECK_FOR_PASSED_ERROR(number_of_characters_v)
	if(string_v.type == ExcelEmpty) return BLANK;
	if(number_of_characters_v.type == ExcelEmpty) return BLANK;
	
	int number_of_characters = (int) number_from(number_of_characters_v);
	CHECK_FOR_CONVERSION_ERROR

	char *string; 
	switch (string_v.type) {
  	  case ExcelString:
  		string = string_v.string;
  		break;
  	  case ExcelNumber:
		  string = malloc(20);
		  if(string == 0) {
			  printf("Out of memory");
			  exit(-1);
		  }
		  snprintf(string,20,"%f",string_v.number);
		  break;
	  case ExcelBoolean:
	  	if(string_v.number == true) {
	  		string = "TRUE";
		} else {
			string = "FALSE";
		}
		break;
	  case ExcelEmpty:	  	 
  	  case ExcelError:
  	  case ExcelRange:
		return string_v;
	}
	
	char *left_string = malloc(number_of_characters+1);
	if(left_string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}	
	memcpy(left_string,string,number_of_characters);
	left_string[number_of_characters] = '\0';
	return new_excel_string(left_string);
}

static ExcelValue left_1(ExcelValue string_v) {
	return left(string_v, ONE);
}

static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error) {
	if(value.type == ExcelError) return value_if_error;
	return value;
}

static ExcelValue more_than(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number <= b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) <= 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue more_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number < b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) < 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}


static ExcelValue less_than(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number >= b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) >= 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue less_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	switch (a_v.type) {
  	  case ExcelNumber:
	  case ExcelBoolean: 
	  case ExcelEmpty:
		if((b_v.type == ExcelNumber) || (b_v.type == ExcelBoolean) || (b_v.type == ExcelEmpty)) {
			if(a_v.number > b_v.number) return FALSE;
			return TRUE;
		} 
		return FALSE;
	  case ExcelString:
	  	if(b_v.type == ExcelString) {
		  	if(strcasecmp(a_v.string,b_v.string) > 0 ) return FALSE;
			return TRUE;	  		
		}
		return FALSE;
  	  case ExcelError:
		return a_v;
  	  case ExcelRange:
  		return NA;
  }
  return FALSE;
}

static ExcelValue subtract(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a - b);
}

static ExcelValue multiply(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a * b);
}

static ExcelValue sum(int array_size, ExcelValue *array) {
	double total = 0;
	int i;
	for(i=0;i<array_size;i++) {
    switch(array[i].type) {
      case ExcelNumber:
        total += array[i].number;
        break;
      case ExcelRange:
        total += number_from(sum( array[i].rows * array[i].columns, array[i].array ));
        break;
      case ExcelError:
        return array[i];
        break;
      default:
        break;
    }
	}
	return new_excel_number(total);
}

static ExcelValue max(int number_of_arguments, ExcelValue *arguments) {
	double biggest_number_found;
	int any_number_found = 0;
	int i;
	ExcelValue current_excel_value;
	
	for(i=0;i<number_of_arguments;i++) {
		current_excel_value = arguments[i];
		if(current_excel_value.type == ExcelNumber) {
			if(!any_number_found) {
				any_number_found = 1;
				biggest_number_found = current_excel_value.number;
			}
			if(current_excel_value.number > biggest_number_found) biggest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelRange) {
			current_excel_value = max( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(current_excel_value.type == ExcelError) return current_excel_value;
			if(current_excel_value.type == ExcelNumber)
				if(!any_number_found) {
					any_number_found = 1;
					biggest_number_found = current_excel_value.number;
				}
				if(current_excel_value.number > biggest_number_found) biggest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelError) {
			return current_excel_value;
		}
	}
	if(!any_number_found) {
		any_number_found = 1;
		biggest_number_found = 0;
	}
	return new_excel_number(biggest_number_found);	
}

static ExcelValue min(int number_of_arguments, ExcelValue *arguments) {
	double smallest_number_found = 0;
	int any_number_found = 0;
	int i;
	ExcelValue current_excel_value;
	
	for(i=0;i<number_of_arguments;i++) {
		current_excel_value = arguments[i];
		if(current_excel_value.type == ExcelNumber) {
			if(!any_number_found) {
				any_number_found = 1;
				smallest_number_found = current_excel_value.number;
			}
			if(current_excel_value.number < smallest_number_found) smallest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelRange) {
			current_excel_value = min( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(current_excel_value.type == ExcelError) return current_excel_value;
			if(current_excel_value.type == ExcelNumber)
				if(!any_number_found) {
					any_number_found = 1;
					smallest_number_found = current_excel_value.number;
				}
				if(current_excel_value.number < smallest_number_found) smallest_number_found = current_excel_value.number; 				
		} else if(current_excel_value.type == ExcelError) {
			return current_excel_value;
		}
	}
	if(!any_number_found) {
		any_number_found = 1;
		smallest_number_found = 0;
	}
	return new_excel_number(smallest_number_found);	
}

static ExcelValue mod(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
		
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return new_excel_number(fmod(a,b));
}

static ExcelValue negative(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(-a);
}

static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v) {
	CHECK_FOR_PASSED_ERROR(rate_v)
	CHECK_FOR_PASSED_ERROR(number_of_periods_v)
	CHECK_FOR_PASSED_ERROR(present_value_v)
		
	NUMBER(rate_v,rate)
	NUMBER(number_of_periods_v,number_of_periods)
	NUMBER(present_value_v,present_value)
	CHECK_FOR_CONVERSION_ERROR
	
	if(rate == 0) return new_excel_number(-(present_value / number_of_periods));
	return new_excel_number(-present_value*(rate*(pow((1+rate),number_of_periods)))/((pow((1+rate),number_of_periods))-1));
}

static ExcelValue power(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
		
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(pow(a,b));
}

static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	
	return new_excel_number( round(number * multiple) / multiple );
}

static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	
	return new_excel_number( trunc(number * multiple) / multiple );	
}

static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)
		
	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR
		
	double multiple = pow(10,decimal_places);
	if(number < 0) return new_excel_number( floor(number * multiple) / multiple );
	return new_excel_number( ceil(number * multiple) / multiple );	
}

static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments) {
	int allocated_length = 100;
	int used_length = 0;
	char *string = malloc(allocated_length);
	if(string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}		
	char *current_string;
	int current_string_length;
	ExcelValue current_v;
	int i;
	for(i=0;i<number_of_arguments;i++) {
		current_v = (ExcelValue) arguments[i];
		switch (current_v.type) {
  	  case ExcelString:
	  		current_string = current_v.string;
	  		break;
  	  case ExcelNumber:
			  current_string = malloc(20);
		  	if(current_string == 0) {
		  	  printf("Out of memory");
		  	  exit(-1);
		  	}				  
			  snprintf(current_string,20,"%g",current_v.number);
			  break;
		  case ExcelBoolean:
		  	if(current_v.number == true) {
		  		current_string = "TRUE";
  			} else {
  				current_string = "FALSE";
  			}
        break;
		  case ExcelEmpty:
        current_string = "";
        break;
      case ExcelError:
        return current_v;
	  	case ExcelRange:
        return VALUE;
		}
		current_string_length = strlen(current_string);
		if( (used_length + current_string_length + 1) > allocated_length) {
			allocated_length += 100;
			string = realloc(string,allocated_length);
		}
		memcpy(string + used_length, current_string, current_string_length);
		used_length = used_length + current_string_length;
	}
	string = realloc(string,used_length+1);
	return new_excel_string(string);
}

static ExcelValue subtotal(ExcelValue subtotal_type_v, int number_of_arguments, ExcelValue *arguments) {
  CHECK_FOR_PASSED_ERROR(subtotal_type_v)
  NUMBER(subtotal_type_v,subtotal_type)
  CHECK_FOR_CONVERSION_ERROR
      
  switch((int) subtotal_type) {
    case 1:
    case 101:
      return average(number_of_arguments,arguments);
      break;
    case 2:
    case 102:
      return count(number_of_arguments,arguments);
      break;
    case 3:
    case 103:
      return counta(number_of_arguments,arguments);
      break;
    case 9:
    case 109:
      return sum(number_of_arguments,arguments);
      break;
    default:
      return VALUE;
      break;
  }
}

static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments) {
  // First, set up the sum_range
  CHECK_FOR_PASSED_ERROR(sum_range_v);

  // Set up the sum range
  ExcelValue *sum_range;
  int sum_range_rows, sum_range_columns;
  
  if(sum_range_v.type == ExcelRange) {
    sum_range = sum_range_v.array;
    sum_range_rows = sum_range_v.rows;
    sum_range_columns = sum_range_v.columns;
  } else {
    sum_range = (ExcelValue*) new_excel_value_array(1);
	sum_range[0] = sum_range_v;
    sum_range_rows = 1;
    sum_range_columns = 1;
  }
  
  // Then go through and set up the check ranges
  if(number_of_arguments % 2 != 0) return VALUE;
  int number_of_criteria = number_of_arguments / 2;
  ExcelValue *criteria_range =  (ExcelValue*) new_excel_value_array(number_of_criteria);
  ExcelValue current_value;
  int i;
  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[i*2];
    if(current_value.type == ExcelRange) {
      criteria_range[i] = current_value;
      if(current_value.rows != sum_range_rows) return VALUE;
      if(current_value.columns != sum_range_columns) return VALUE;
    } else {
      if(sum_range_rows != 1) return VALUE;
      if(sum_range_columns != 1) return VALUE;
      ExcelValue *tmp_array2 =  (ExcelValue*) new_excel_value_array(1);
      tmp_array2[0] = current_value;
      criteria_range[i] =  new_excel_range(tmp_array2,1,1);
    }
  }
  
  // Now go through and set up the criteria
  ExcelComparison *criteria =  malloc(sizeof(ExcelComparison)*number_of_criteria);
  if(criteria == 0) {
	  printf("Out of memory\n");
	  exit(-1);
  }
  char *s;
  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[(i*2)+1];
    
    if(current_value.type == ExcelString) {
      s = current_value.string;
      if(s[0] == '<') {
        if( s[1] == '>') {
          criteria[i].type = NotEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else if(s[1] == '=') {
          criteria[i].type = LessThanOrEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else {
          criteria[i].type = LessThan;
          criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));
        }
      } else if(s[0] == '>') {
        if(s[1] == '=') {
          criteria[i].type = MoreThanOrEqual;
          criteria[i].comparator = new_excel_string(strndup(s+2,strlen(s)-2));
        } else {
          criteria[i].type = MoreThan;
          criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));
        }
      } else if(s[0] == '=') {
        criteria[i].type = Equal;
        criteria[i].comparator = new_excel_string(strndup(s+1,strlen(s)-1));          
      } else {
        criteria[i].type = Equal;
        criteria[i].comparator = current_value;          
      }
    } else {
      criteria[i].type = Equal;
      criteria[i].comparator = current_value;
    }
  }
  
  double total = 0;
  int size = sum_range_columns * sum_range_rows;
  int j;
  int passed = 0;
  ExcelValue value_to_be_checked;
  ExcelComparison comparison;
  ExcelValue comparator;
  double number;
  // For each cell in the sum range
  for(j=0; j < size; j++ ) {
    passed = 1;
    for(i=0; i < number_of_criteria; i++) {
      value_to_be_checked = ((ExcelValue *) ((ExcelValue) criteria_range[i]).array)[j];
      comparison = criteria[i];
      comparator = comparison.comparator;
      
      switch(value_to_be_checked.type) {
        case ExcelError: // Errors match only errors
          if(comparison.type != Equal) passed = 0;
          if(comparator.type != ExcelError) passed = 0;
          if(value_to_be_checked.number != comparator.number) passed = 0;
          break;
        case ExcelBoolean: // Booleans match only booleans (FIXME: I think?)
          if(comparison.type != Equal) passed = 0;
          if(comparator.type != ExcelBoolean ) passed = 0;
          if(value_to_be_checked.number != comparator.number) passed = 0;
          break;
        case ExcelEmpty:
          // if(comparator.type == ExcelEmpty) break; // FIXME: Huh? In excel blank doesn't match blank?!
          if(comparator.type != ExcelString) {
            passed = 0;
            break;
          } else {
            if(strlen(comparator.string) != 0) passed = 0; // Empty strings match blanks.
            break;
          }
          break;
        case ExcelNumber:
          if(comparator.type == ExcelNumber) {
            number = comparator.number;
          } else if(comparator.type == ExcelString) {
            number = number_from(comparator);
            if(conversion_error == 1) {
              conversion_error = 0;
              passed = 0;
              break;
            }
          } else {
            passed = 0;
            break;
          }
          switch(comparison.type) {
            case Equal:
              if(value_to_be_checked.number != number) passed = 0;
              break;
            case LessThan:
              if(value_to_be_checked.number >= number) passed = 0;
              break;            
            case LessThanOrEqual:
              if(value_to_be_checked.number > number) passed = 0;
              break;                        
            case NotEqual:
              if(value_to_be_checked.number == number) passed = 0;
              break;            
            case MoreThanOrEqual:
              if(value_to_be_checked.number < number) passed = 0;
              break;            
            case MoreThan:
              if(value_to_be_checked.number <= number) passed = 0;
              break;
          }
          break;
        case ExcelString:
          // First case, the comparator is a number, simplification is that it can only be equal
          if(comparator.type == ExcelNumber) {
            if(comparison.type != Equal) {
              printf("This shouldn't be possible?");
              passed = 0;
              break;
            }
            number = number_from(value_to_be_checked);
            if(conversion_error == 1) {
              conversion_error = 0;
              passed = 0;
              break;
            }
            if(number != comparator.number) {
              passed = 0;
              break;
            } else {
              break;
            }
          // Second case, the comparator is also a string, so need to be able to do full range of tests
          } else if(comparator.type == ExcelString) {
            switch(comparison.type) {
              case Equal:
                if(excel_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;
              case LessThan:
                if(less_than(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case LessThanOrEqual:
                if(less_than_or_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;                        
              case NotEqual:
                if(not_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case MoreThanOrEqual:
                if(more_than_or_equal(value_to_be_checked,comparator).number == 0) passed = 0;
                break;            
              case MoreThan:
                if(more_than(value_to_be_checked,comparator).number == 0) passed = 0;
                break;
              }
          } else {
            passed = 0;
            break;
          }
          break;
        case ExcelRange:
          return VALUE;            
      }
      if(passed == 0) break;
    }
    if(passed == 1) {
      current_value = sum_range[j];
      if(current_value.type == ExcelError) {
        return current_value;
      } else if(current_value.type == ExcelNumber) {
        total += current_value.number;
      }
    }
  }
  return new_excel_number(total);
}

static ExcelValue sumif(ExcelValue check_range_v, ExcelValue criteria_v, ExcelValue sum_range_v ) {
	ExcelValue tmp_array_sumif[] = {check_range_v, criteria_v};
	return sumifs(sum_range_v,2,tmp_array_sumif);
}

static ExcelValue sumif_2(ExcelValue check_range_v, ExcelValue criteria_v) {
	ExcelValue tmp_array_sumif2[] = {check_range_v, criteria_v};
	return sumifs(check_range_v,2,tmp_array_sumif2);
}

static ExcelValue sumproduct(int number_of_arguments, ExcelValue *arguments) {
  if(number_of_arguments <1) return VALUE;
  
  int a;
  int i;
  int j;
  int rows;
  int columns;
  ExcelValue current_value;
  ExcelValue **ranges = malloc(sizeof(ExcelValue *)*number_of_arguments);
  if(ranges == 0) {
	  printf("Out of memory\n");
	  exit(-1);
  }
  double product = 1;
  double sum = 0;
  
  // Find out dimensions of first argument
  if(arguments[0].type == ExcelRange) {
    rows = arguments[0].rows;
    columns = arguments[0].columns;
  } else {
    rows = 1;
    columns = 1;
  }
  // Extract arrays from each of the given ranges, checking for errors and bounds as we go
  for(a=0;a<number_of_arguments;a++) {
    current_value = arguments[a];
    switch(current_value.type) {
      case ExcelRange:
        if(current_value.rows != rows || current_value.columns != columns) return VALUE;
        ranges[a] = current_value.array;
        break;
      case ExcelError:
        return current_value;
        break;
      case ExcelEmpty:
        return VALUE;
        break;
      default:
        if(rows != 1 && columns !=1) return VALUE;
        ranges[a] = (ExcelValue*) new_excel_value_array(1);
        ranges[a][0] = arguments[a];
        break;
    }
  }
  	
	for(i=0;i<rows;i++) {
		for(j=0;j<columns;j++) {
			product = 1;
			for(a=0;a<number_of_arguments;a++) {
				current_value = ranges[a][(i*columns)+j];
				if(current_value.type == ExcelNumber) {
					product *= current_value.number;
				} else {
					product *= 0;
				}
			}
			sum += product;
		}
	}
  	return new_excel_number(sum);
}

static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v) {
  return vlookup(lookup_value_v,lookup_table_v,column_number_v,TRUE);
}

static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v) {
  CHECK_FOR_PASSED_ERROR(lookup_value_v)
  CHECK_FOR_PASSED_ERROR(lookup_table_v)
  CHECK_FOR_PASSED_ERROR(column_number_v)
  CHECK_FOR_PASSED_ERROR(match_type_v)

  if(lookup_value_v.type == ExcelEmpty) return NA;
  if(lookup_table_v.type != ExcelRange) return NA;
  if(column_number_v.type != ExcelNumber) return NA;
  if(match_type_v.type != ExcelBoolean) return NA;
    
  int i;
  int last_good_match = 0;
  int rows = lookup_table_v.rows;
  int columns = lookup_table_v.columns;
  ExcelValue *array = lookup_table_v.array;
  ExcelValue possible_match_v;
  
  if(column_number_v.number > columns) return REF;
  if(column_number_v.number < 1) return VALUE;
  
  if(match_type_v.number == false) { // Exact match required
    for(i=0; i< rows; i++) {
      possible_match_v = array[i*columns];
      if(excel_equal(lookup_value_v,possible_match_v).number == true) {
        return array[(i*columns)+(((int) column_number_v.number) - 1)];
      }
    }
    return NA;
  } else { // Highest value that is less than or equal
    for(i=0; i< rows; i++) {
      possible_match_v = array[i*columns];
      if(lookup_value_v.type != possible_match_v.type) continue;
      if(more_than(possible_match_v,lookup_value_v).number == true) {
        if(i == 0) return NA;
        return array[((i-1)*columns)+(((int) column_number_v.number) - 1)];
      } else {
        last_good_match = i;
      }
    }
    return array[(last_good_match*columns)+(((int) column_number_v.number) - 1)];   
  }
  return NA;
}



int test_functions() {
	// Test ABS
	assert(excel_abs(ONE).number == 1);
	assert(excel_abs(new_excel_number(-1)).number == 1);
	assert(excel_abs(VALUE).type == ExcelError);
	
	// Test ADD
	assert(add(ONE,new_excel_number(-2.5)).number == -1.5);
	assert(add(ONE,VALUE).type == ExcelError);
	
	// Test AND
	ExcelValue true_array1[] = { TRUE, new_excel_number(10)};
	ExcelValue true_array2[] = { ONE };
	ExcelValue false_array1[] = { FALSE, new_excel_number(10)};
	ExcelValue false_array2[] = { TRUE, new_excel_number(0)};
	// ExcelValue error_array1[] = { new_excel_number(10)}; // Not implemented
	ExcelValue error_array2[] = { TRUE, NA};
	assert(excel_and(2,true_array1).number == 1);
	assert(excel_and(1,true_array2).number == 1);
	assert(excel_and(2,false_array1).number == 0);
	assert(excel_and(2,false_array2).number == 0);
	// assert(excel_and(1,error_array1).type == ExcelError); // Not implemented
	assert(excel_and(2,error_array2).type == ExcelError);
	
	// Test AVERAGE
	ExcelValue array1[] = { new_excel_number(10), new_excel_number(5), TRUE, FALSE};
	ExcelValue array1_v = new_excel_range(array1,2,2);
	ExcelValue array2[] = { array1_v, new_excel_number(9), new_excel_string("Hello")};
	ExcelValue array3[] = { array1_v, new_excel_number(9), new_excel_string("Hello"), VALUE};
	assert(average(4, array1).number == 7.5);
	assert(average(3, array2).number == 8);
	assert(average(4, array3).type == ExcelError);
	
	// Test CHOOSE
	assert(choose(ONE,4,array1).number == 10);
	assert(choose(new_excel_number(4),4,array1).type == ExcelBoolean);
	assert(choose(new_excel_number(0),4,array1).type == ExcelError);
	assert(choose(new_excel_number(5),4,array1).type == ExcelError);
	assert(choose(ONE,4,array3).type == ExcelError);	
	
	// Test COUNT
	assert(count(4,array1).number == 2);
	assert(count(3,array2).number == 3);
	assert(count(4,array3).number == 3);
	
	// Test COUNTA
	ExcelValue count_a_test_array_1[] = { new_excel_number(10), new_excel_number(5), TRUE, FALSE, new_excel_string("Hello"), VALUE, BLANK};
  ExcelValue count_a_test_array_1_v = new_excel_range(count_a_test_array_1,7,1);
  ExcelValue count_a_test_array_2[] = {new_excel_string("Bye"),count_a_test_array_1_v};
	assert(counta(7, count_a_test_array_1).number == 6);
  assert(counta(2, count_a_test_array_2).number == 7);
	
	// Test divide
	assert(divide(new_excel_number(12.4),new_excel_number(3.2)).number == 3.875);
	assert(divide(new_excel_number(12.4),new_excel_number(0)).type == ExcelError);
	
	// Test excel_equal
	assert(excel_equal(new_excel_number(1.2),new_excel_number(3.4)).type == ExcelBoolean);
	assert(excel_equal(new_excel_number(1.2),new_excel_number(3.4)).number == false);
	assert(excel_equal(new_excel_number(1.2),new_excel_number(1.2)).number == true);
	assert(excel_equal(new_excel_string("hello"), new_excel_string("HELLO")).number == true);
	assert(excel_equal(new_excel_string("hello world"), new_excel_string("HELLO")).number == false);
	assert(excel_equal(new_excel_string("1"), ONE).number == false);
	assert(excel_equal(DIV0, ONE).type == ExcelError);

	// Test not_equal
	assert(not_equal(new_excel_number(1.2),new_excel_number(3.4)).type == ExcelBoolean);
	assert(not_equal(new_excel_number(1.2),new_excel_number(3.4)).number == true);
	assert(not_equal(new_excel_number(1.2),new_excel_number(1.2)).number == false);
	assert(not_equal(new_excel_string("hello"), new_excel_string("HELLO")).number == false);
	assert(not_equal(new_excel_string("hello world"), new_excel_string("HELLO")).number == true);
	assert(not_equal(new_excel_string("1"), ONE).number == true);
	assert(not_equal(DIV0, ONE).type == ExcelError);
	
	// Test excel_if
	// Two argument version
	assert(excel_if_2(TRUE,new_excel_number(10)).type == ExcelNumber);
	assert(excel_if_2(TRUE,new_excel_number(10)).number == 10);
	assert(excel_if_2(FALSE,new_excel_number(10)).type == ExcelBoolean);
	assert(excel_if_2(FALSE,new_excel_number(10)).number == false);
	assert(excel_if_2(NA,new_excel_number(10)).type == ExcelError);
	// Three argument version
	assert(excel_if(TRUE,new_excel_number(10),new_excel_number(20)).type == ExcelNumber);
	assert(excel_if(TRUE,new_excel_number(10),new_excel_number(20)).number == 10);
	assert(excel_if(FALSE,new_excel_number(10),new_excel_number(20)).type == ExcelNumber);
	assert(excel_if(FALSE,new_excel_number(10),new_excel_number(20)).number == 20);
	assert(excel_if(NA,new_excel_number(10),new_excel_number(20)).type == ExcelError);
	
	// Test excel_match
	ExcelValue excel_match_array_1[] = { new_excel_number(10), new_excel_number(100) };
	ExcelValue excel_match_array_1_v = new_excel_range(excel_match_array_1,1,2);
	ExcelValue excel_match_array_2[] = { new_excel_string("Pear"), new_excel_string("Bear"), new_excel_string("Apple") };
	ExcelValue excel_match_array_2_v = new_excel_range(excel_match_array_2,3,1);
	ExcelValue excel_match_array_4[] = { ONE, BLANK, new_excel_number(0) };
	ExcelValue excel_match_array_4_v = new_excel_range(excel_match_array_4,1,3);
	ExcelValue excel_match_array_5[] = { ONE, new_excel_number(0), BLANK };
	ExcelValue excel_match_array_5_v = new_excel_range(excel_match_array_5,1,3);
	
	// Two argument version
	assert(excel_match_2(new_excel_number(10),excel_match_array_1_v).number == 1);
	assert(excel_match_2(new_excel_number(100),excel_match_array_1_v).number == 2);
	assert(excel_match_2(new_excel_number(1000),excel_match_array_1_v).type == ExcelError);
    assert(excel_match_2(new_excel_number(0), excel_match_array_4_v).number == 2);
    assert(excel_match_2(BLANK, excel_match_array_5_v).number == 2);

	// Three argument version	
    assert(excel_match(new_excel_number(10.0), excel_match_array_1_v, new_excel_number(0) ).number == 1);
    assert(excel_match(new_excel_number(100.0), excel_match_array_1_v, new_excel_number(0) ).number == 2);
    assert(excel_match(new_excel_number(1000.0), excel_match_array_1_v, new_excel_number(0) ).type == ExcelError);
    assert(excel_match(new_excel_string("bEAr"), excel_match_array_2_v, new_excel_number(0) ).number == 2);
    assert(excel_match(new_excel_number(1000.0), excel_match_array_1_v, ONE ).number == 2);
    assert(excel_match(new_excel_number(1.0), excel_match_array_1_v, ONE ).type == ExcelError);
    assert(excel_match(new_excel_string("Care"), excel_match_array_2_v, new_excel_number(-1) ).number == 1  );
    assert(excel_match(new_excel_string("Zebra"), excel_match_array_2_v, new_excel_number(-1) ).type == ExcelError);
    assert(excel_match(new_excel_string("a"), excel_match_array_2_v, new_excel_number(-1) ).number == 2);
	
	// When not given a range
    assert(excel_match(new_excel_number(10.0), new_excel_number(10), new_excel_number(0.0)).number == 1);
    assert(excel_match(new_excel_number(20.0), new_excel_number(10), new_excel_number(0.0)).type == ExcelError);
    assert(excel_match(new_excel_number(10.0), excel_match_array_1_v, BLANK).number == 1);

	// Test more than on
	// .. numbers
    assert(more_than(ONE,new_excel_number(2)).number == false);
    assert(more_than(ONE,ONE).number == false);
    assert(more_than(ONE,new_excel_number(0)).number == true);
	// .. booleans
    assert(more_than(FALSE,FALSE).number == false);
    assert(more_than(FALSE,TRUE).number == false);
    assert(more_than(TRUE,FALSE).number == true);
    assert(more_than(TRUE,TRUE).number == false);
	// ..strings
    assert(more_than(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == true);		
    assert(more_than(new_excel_string("HELLO"),new_excel_string("world")).number == false);
    assert(more_than(new_excel_string("HELLO"),new_excel_string("hello")).number == false);
	// ..blanks
    assert(more_than(BLANK,ONE).number == false);
    assert(more_than(BLANK,new_excel_number(-1)).number == true);
    assert(more_than(ONE,BLANK).number == true);
    assert(more_than(new_excel_number(-1),BLANK).number == false);

	// Test less than on
	// .. numbers
    assert(less_than(ONE,new_excel_number(2)).number == true);
    assert(less_than(ONE,ONE).number == false);
    assert(less_than(ONE,new_excel_number(0)).number == false);
	// .. booleans
    assert(less_than(FALSE,FALSE).number == false);
    assert(less_than(FALSE,TRUE).number == true);
    assert(less_than(TRUE,FALSE).number == false);
    assert(less_than(TRUE,TRUE).number == false);
	// ..strings
    assert(less_than(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == false);		
    assert(less_than(new_excel_string("HELLO"),new_excel_string("world")).number == true);
    assert(less_than(new_excel_string("HELLO"),new_excel_string("hello")).number == false);
	// ..blanks
    assert(less_than(BLANK,ONE).number == true);
    assert(less_than(BLANK,new_excel_number(-1)).number == false);
    assert(less_than(ONE,BLANK).number == false);
    assert(less_than(new_excel_number(-1),BLANK).number == true);

	// Test FIND function
	// ... should find the first occurrence of one string in another, returning :value if the string doesn't match
	assert(find_2(new_excel_string("one"),new_excel_string("onetwothree")).number == 1);
	assert(find_2(new_excel_string("one"),new_excel_string("twoonethree")).number == 4);
	assert(find_2(new_excel_string("one"),new_excel_string("twoonthree")).type == ExcelError);
    // ... should find the first occurrence of one string in another after a given index, returning :value if the string doesn't match
	assert(find(new_excel_string("one"),new_excel_string("onetwothree"),ONE).number == 1);
	assert(find(new_excel_string("one"),new_excel_string("twoonethree"),new_excel_number(5)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(2)).number == 4);
    // ... should be possible for the start_num to be a string, if that string converts to a number
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_string("2")).number == 4);
    // ... should return a :value error when given anything but a number as the third argument
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_string("a")).type == ExcelError);
    // ... should return a :value error when given a third argument that is less than 1 or greater than the length of the string
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(0)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(-1)).type == ExcelError);
	assert(find(new_excel_string("one"),new_excel_string("oneone"),new_excel_number(7)).type == ExcelError);
	// ... BLANK in the first argument matches any character
	assert(find_2(BLANK,new_excel_string("abcdefg")).number == 1);
	assert(find(BLANK,new_excel_string("abcdefg"),new_excel_number(4)).number == 4);
    // ... should treat BLANK in the second argument as an empty string
	assert(find_2(BLANK,BLANK).number == 1);
	assert(find_2(new_excel_string("a"),BLANK).type == ExcelError);
	// ... should return an error if any argument is an error
	assert(find(new_excel_string("one"),new_excel_string("onetwothree"),NA).type == ExcelError);
	assert(find(new_excel_string("one"),NA,ONE).type == ExcelError);
	assert(find(NA,new_excel_string("onetwothree"),ONE).type == ExcelError);
	
	// Test the IFERROR function
    assert(iferror(new_excel_string("ok"),ONE).type == ExcelString);
	assert(iferror(VALUE,ONE).type == ExcelNumber);		
	
	// Test the INDEX function
	ExcelValue index_array_1[] = { new_excel_number(10), new_excel_number(20), BLANK };
	ExcelValue index_array_1_v_column = new_excel_range(index_array_1,3,1);
	ExcelValue index_array_1_v_row = new_excel_range(index_array_1,1,3);
	ExcelValue index_array_2[] = { BLANK, ONE, new_excel_number(10), new_excel_number(11), new_excel_number(100), new_excel_number(101) };
	ExcelValue index_array_2_v = new_excel_range(index_array_2,3,2);
	// ... if given one argument should return the value at that offset in the range
	assert(excel_index_2(index_array_1_v_column,new_excel_number(2.0)).number == 20);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(2.0)).number == 20);
	// ... but not if the range is not a single row or single column
	assert(excel_index_2(index_array_2_v,new_excel_number(2.0)).type == ExcelError);
    // ... it should return the value in the array at position row_number, column_number
	assert(excel_index(new_excel_number(10),ONE,ONE).number == 10);
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(2.0)).number == 1);
	assert(excel_index(index_array_2_v,new_excel_number(2.0),new_excel_number(1.0)).number == 10);
	assert(excel_index(index_array_2_v,new_excel_number(3.0),new_excel_number(1.0)).number == 100);
	assert(excel_index(index_array_2_v,new_excel_number(3.0),new_excel_number(3.0)).type == ExcelError);
	// ... it should return ZERO not blank, if a blank cell is picked
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(1.0)).type == ExcelNumber);
	assert(excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(1.0)).number == 0);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(3.0)).type == ExcelNumber);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(3.0)).number == 0);
	// ... it should return the whole row if given a zero column number
	ExcelValue index_result_1_v = excel_index(index_array_2_v,new_excel_number(1.0),new_excel_number(0.0));
	assert(index_result_1_v.type == ExcelRange);
	assert(index_result_1_v.rows == 1);
	assert(index_result_1_v.columns == 2);
	ExcelValue *index_result_1_a = index_result_1_v.array;
	assert(index_result_1_a[0].number == 0);
	assert(index_result_1_a[1].number == 1);
	// ... it should return the whole column if given a zero row number
	ExcelValue index_result_2_v = excel_index(index_array_2_v,new_excel_number(0),new_excel_number(1.0));
	assert(index_result_2_v.type == ExcelRange);
	assert(index_result_2_v.rows == 3);
	assert(index_result_2_v.columns == 1);
	ExcelValue *index_result_2_a = index_result_2_v.array;
	assert(index_result_2_a[0].number == 0);
	assert(index_result_2_a[1].number == 10);
	assert(index_result_2_a[2].number == 100);
    // ... it should return a :ref error when given arguments outside array range
	assert(excel_index_2(index_array_1_v_row,new_excel_number(-1)).type == ExcelError);
	assert(excel_index_2(index_array_1_v_row,new_excel_number(4)).type == ExcelError);
    // ... it should treat BLANK as zero if given as a required row or column number
	assert(excel_index(index_array_2_v,new_excel_number(1.0),BLANK).type == ExcelRange);
	assert(excel_index(index_array_2_v,BLANK,new_excel_number(2.0)).type == ExcelRange);
    // ... it should return an error if an argument is an error
	assert(excel_index(NA,NA,NA).type == ExcelError);
	
	// LEFT(string,[characters])
	// ... should return the left n characters from a string
    assert(strcmp(left_1(new_excel_string("ONE")).string,"O") == 0);
    assert(strcmp(left(new_excel_string("ONE"),ONE).string,"O") == 0);
    assert(strcmp(left(new_excel_string("ONE"),new_excel_number(3)).string,"ONE") == 0);
	// ... should turn numbers into strings before processing
	assert(strcmp(left(new_excel_number(1.31e12),new_excel_number(3)).string, "131") == 0);
	// ... should turn booleans into the words TRUE and FALSE before processing
    assert(strcmp(left(TRUE,new_excel_number(3)).string,"TRU") == 0);
	assert(strcmp(left(FALSE,new_excel_number(3)).string,"FAL") == 0);
	// ... should return BLANK if given BLANK for either argument
	assert(left(BLANK,new_excel_number(3)).type == ExcelEmpty);
	assert(left(new_excel_string("ONE"),BLANK).type == ExcelEmpty);
	// ... should return an error if an argument is an error
    assert(left_1(NA).type == ExcelError);
    assert(left(new_excel_string("ONE"),NA).type == ExcelError);
	
	// Test less than or equal to
	// .. numbers
    assert(less_than_or_equal(ONE,new_excel_number(2)).number == true);
    assert(less_than_or_equal(ONE,ONE).number == true);
    assert(less_than_or_equal(ONE,new_excel_number(0)).number == false);
	// .. booleans
    assert(less_than_or_equal(FALSE,FALSE).number == true);
    assert(less_than_or_equal(FALSE,TRUE).number == true);
    assert(less_than_or_equal(TRUE,FALSE).number == false);
    assert(less_than_or_equal(TRUE,TRUE).number == true);
	// ..strings
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == false);		
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("world")).number == true);
    assert(less_than_or_equal(new_excel_string("HELLO"),new_excel_string("hello")).number == true);
	// ..blanks
    assert(less_than_or_equal(BLANK,ONE).number == true);
    assert(less_than_or_equal(BLANK,new_excel_number(-1)).number == false);
    assert(less_than_or_equal(ONE,BLANK).number == false);
    assert(less_than_or_equal(new_excel_number(-1),BLANK).number == true);

	// Test MAX
	assert(max(4, array1).number == 10);
	assert(max(3, array2).number == 10);
	assert(max(4, array3).type == ExcelError);

	// Test MIN
	assert(min(4, array1).number == 5);
	assert(min(3, array2).number == 5);
	assert(min(4, array3).type == ExcelError);

	// Test MOD
    // ... should return the remainder of a number
	assert(mod(new_excel_number(10), new_excel_number(3)).number == 1.0);
	assert(mod(new_excel_number(10), new_excel_number(5)).number == 0.0);
    // ... should be possible for the the arguments to be strings, if they convert to a number
	assert(mod(new_excel_string("3.5"),new_excel_string("2")).number == 1.5);
    // ... should treat BLANK as zero
	assert(mod(BLANK,new_excel_number(10)).number == 0);
	assert(mod(new_excel_number(10),BLANK).type == ExcelError);
	assert(mod(BLANK,BLANK).type == ExcelError);
    // ... should treat true as 1 and FALSE as 0
	assert((mod(new_excel_number(1.1),TRUE).number - 0.1) < 0.001);	
	assert(mod(new_excel_number(1.1),FALSE).type == ExcelError);
	assert(mod(FALSE,new_excel_number(10)).number == 0);
    // ... should return an error when given inappropriate arguments
	assert(mod(new_excel_string("Asdasddf"),new_excel_string("adsfads")).type == ExcelError);
    // ... should return an error if an argument is an error
	assert(mod(new_excel_number(1),VALUE).type == ExcelError);
	assert(mod(VALUE,new_excel_number(1)).type == ExcelError);
	assert(mod(VALUE,VALUE).type == ExcelError);
	
	// Test more than or equal to on
	// .. numbers
    assert(more_than_or_equal(ONE,new_excel_number(2)).number == false);
    assert(more_than_or_equal(ONE,ONE).number == true);
    assert(more_than_or_equal(ONE,new_excel_number(0)).number == true);
	// .. booleans
    assert(more_than_or_equal(FALSE,FALSE).number == true);
    assert(more_than_or_equal(FALSE,TRUE).number == false);
    assert(more_than_or_equal(TRUE,FALSE).number == true);
    assert(more_than_or_equal(TRUE,TRUE).number == true);
	// ..strings
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("Ardvark")).number == true);		
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("world")).number == false);
    assert(more_than_or_equal(new_excel_string("HELLO"),new_excel_string("hello")).number == true);
	// ..blanks
    assert(more_than_or_equal(BLANK,BLANK).number == true);
    assert(more_than_or_equal(BLANK,ONE).number == false);
    assert(more_than_or_equal(BLANK,new_excel_number(-1)).number == true);
    assert(more_than_or_equal(ONE,BLANK).number == true);
    assert(more_than_or_equal(new_excel_number(-1),BLANK).number == false);	
	
	// Test negative
    // ... should return the negative of its arguments
	assert(negative(new_excel_number(1)).number == -1);
	assert(negative(new_excel_number(-1)).number == 1);
    // ... should treat strings that only contain numbers as numbers
	assert(negative(new_excel_string("10")).number == -10);
	assert(negative(new_excel_string("-1.3")).number == 1.3);
    // ... should return an error when given inappropriate arguments
	assert(negative(new_excel_string("Asdasddf")).type == ExcelError);
    // ... should treat BLANK as zero
	assert(negative(BLANK).number == 0);
	
	// Test PMT(rate,number_of_periods,present_value) - optional arguments not yet implemented
    // ... should calculate the monthly payment required for a given principal, interest rate and loan period
    assert((pmt(new_excel_number(0.1),new_excel_number(10),new_excel_number(100)).number - -16.27) < 0.01);
    assert((pmt(new_excel_number(0.0123),new_excel_number(99.1),new_excel_number(123.32)).number - -2.159) < 0.01);
    assert((pmt(new_excel_number(0),new_excel_number(2),new_excel_number(10)).number - -5) < 0.01);

	// Test power
    // ... should return sum of its arguments
	assert(power(new_excel_number(2),new_excel_number(3)).number == 8);
	assert(power(new_excel_number(4.0),new_excel_number(0.5)).number == 2.0);
	
	// Test round
    assert(excel_round(new_excel_number(1.1), new_excel_number(0)).number == 1.0);
    assert(excel_round(new_excel_number(1.5), new_excel_number(0)).number == 2.0);
    assert(excel_round(new_excel_number(1.56),new_excel_number(1)).number == 1.6);
    assert(excel_round(new_excel_number(-1.56),new_excel_number(1)).number == -1.6);

	// Test rounddown
    assert(rounddown(new_excel_number(1.1), new_excel_number(0)).number == 1.0);
    assert(rounddown(new_excel_number(1.5), new_excel_number(0)).number == 1.0);
    assert(rounddown(new_excel_number(1.56),new_excel_number(1)).number == 1.5);
    assert(rounddown(new_excel_number(-1.56),new_excel_number(1)).number == -1.5);	

	// Test roundup
    assert(roundup(new_excel_number(1.1), new_excel_number(0)).number == 2.0);
    assert(roundup(new_excel_number(1.5), new_excel_number(0)).number == 2.0);
    assert(roundup(new_excel_number(1.56),new_excel_number(1)).number == 1.6);
    assert(roundup(new_excel_number(-1.56),new_excel_number(1)).number == -1.6);	
	
	// Test string joining
	ExcelValue string_join_array_1[] = {new_excel_string("Hello "), new_excel_string("world")};
	ExcelValue string_join_array_2[] = {new_excel_string("Hello "), new_excel_string("world"), new_excel_string("!")};
	ExcelValue string_join_array_3[] = {new_excel_string("Top "), new_excel_number(10.0)};
	ExcelValue string_join_array_4[] = {new_excel_string("Top "), new_excel_number(10.5)};	
	ExcelValue string_join_array_5[] = {new_excel_string("Top "), TRUE, FALSE};	
	// ... should return a string by combining its arguments
	// inspect_excel_value(string_join(2, string_join_array_1));
  assert(string_join(2, string_join_array_1).string[6] == 'w');
	// ... should cope with an arbitrary number of arguments
  assert(string_join(3, string_join_array_2).string[11] == '!');
	// ... should convert values to strings as it goes
  assert(string_join(2, string_join_array_3).string[4] == '1');
	// ... should convert integer values into strings without decimal points
  assert(string_join(2, string_join_array_3).string[7] == '\0');
  assert(string_join(2, string_join_array_4).string[7] == '5');
	// ... should convert TRUE and FALSE into strings
  assert(string_join(3,string_join_array_5).string[4] == 'T');
	
  // Test SUBTOTAL function
  ExcelValue subtotal_array_1[] = {new_excel_number(10),new_excel_number(100),BLANK};
  ExcelValue subtotal_array_1_v = new_excel_range(subtotal_array_1,3,1);
  ExcelValue subtotal_array_2[] = {new_excel_number(1),new_excel_string("two"),subtotal_array_1_v};
  
  // new_excel_number(1.0); 
  // inspect_excel_value(new_excel_number(1.0)); 
  // inspect_excel_value(new_excel_range(subtotal_array_2,3,1)); 
  // inspect_excel_value(subtotal(new_excel_number(1.0),3,subtotal_array_2)); 
  
  assert(subtotal(new_excel_number(1.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(new_excel_number(2.0),3,subtotal_array_2).number == 3);
  assert(subtotal(new_excel_number(3.0),7, count_a_test_array_1).number == 6);
  assert(subtotal(new_excel_number(3.0),3,subtotal_array_2).number == 4);
  assert(subtotal(new_excel_number(9.0),3,subtotal_array_2).number == 111);
  assert(subtotal(new_excel_number(101.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(new_excel_number(102.0),3,subtotal_array_2).number == 3);
  assert(subtotal(new_excel_number(103.0),3,subtotal_array_2).number == 4);
  assert(subtotal(new_excel_number(109.0),3,subtotal_array_2).number == 111);
  
  // Test SUMIFS function
  ExcelValue sumifs_array_1[] = {new_excel_number(10),new_excel_number(100),BLANK};
  ExcelValue sumifs_array_1_v = new_excel_range(sumifs_array_1,3,1);
  ExcelValue sumifs_array_2[] = {new_excel_string("pear"),new_excel_string("bear"),new_excel_string("apple")};
  ExcelValue sumifs_array_2_v = new_excel_range(sumifs_array_2,3,1);
  ExcelValue sumifs_array_3[] = {new_excel_number(1),new_excel_number(2),new_excel_number(3),new_excel_number(4),new_excel_number(5),new_excel_number(5)};
  ExcelValue sumifs_array_3_v = new_excel_range(sumifs_array_3,6,1);
  ExcelValue sumifs_array_4[] = {new_excel_string("CO2"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CO2")};
  ExcelValue sumifs_array_4_v = new_excel_range(sumifs_array_4,6,1);
  ExcelValue sumifs_array_5[] = {new_excel_string("1A"),new_excel_string("1A"),new_excel_string("1A"),new_excel_number(4),new_excel_number(4),new_excel_number(5)};
  ExcelValue sumifs_array_5_v = new_excel_range(sumifs_array_5,6,1);
  
  // ... should only sum values that meet all of the criteria
  ExcelValue sumifs_array_6[] = { sumifs_array_1_v, new_excel_number(10), sumifs_array_2_v, new_excel_string("Bear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_6).number == 0.0);
  
  ExcelValue sumifs_array_7[] = { sumifs_array_1_v, new_excel_number(10), sumifs_array_2_v, new_excel_string("Pear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_7).number == 10.0);
  
  // ... should work when single cells are given where ranges expected
  ExcelValue sumifs_array_8[] = { new_excel_string("CAR"), new_excel_string("CAR"), new_excel_string("FCV"), new_excel_string("FCV")};
  assert(sumifs(new_excel_number(0.143897265452564), 4, sumifs_array_8).number == 0.143897265452564);

  // ... should match numbers with strings that contain numbers
  ExcelValue sumifs_array_9[] = { new_excel_number(10), new_excel_string("10.0")};
  assert(sumifs(new_excel_number(100),2,sumifs_array_9).number == 100);
  
  ExcelValue sumifs_array_10[] = { sumifs_array_4_v, new_excel_string("CO2"), sumifs_array_5_v, new_excel_number(2)};
  assert(sumifs(sumifs_array_3_v,4, sumifs_array_10).number == 0);
  
  // ... should match with strings that contain criteria
  ExcelValue sumifs_array_10a[] = { sumifs_array_3_v, new_excel_string("=5")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10a).number == 10);

  ExcelValue sumifs_array_10b[] = { sumifs_array_3_v, new_excel_string("<>3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10b).number == 17);

  ExcelValue sumifs_array_10c[] = { sumifs_array_3_v, new_excel_string("<3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10c).number == 3);
  
  ExcelValue sumifs_array_10d[] = { sumifs_array_3_v, new_excel_string("<=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10d).number == 6);

  ExcelValue sumifs_array_10e[] = { sumifs_array_3_v, new_excel_string(">3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10e).number == 14);

  ExcelValue sumifs_array_10f[] = { sumifs_array_3_v, new_excel_string(">=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10f).number == 17);
  
  // ... should treat BLANK as an empty string when in the check_range, but not in the criteria
  ExcelValue sumifs_array_11[] = { BLANK, new_excel_number(20)};
  assert(sumifs(new_excel_number(100),2,sumifs_array_11).number == 0);
  
  ExcelValue sumifs_array_12[] = {BLANK, new_excel_string("")};
  assert(sumifs(new_excel_number(100),2,sumifs_array_12).number == 100);
  
  ExcelValue sumifs_array_13[] = {BLANK, BLANK};
  assert(sumifs(new_excel_number(100),2,sumifs_array_13).number == 0);
    
  // ... should return an error if range argument is an error
  assert(sumifs(REF,2,sumifs_array_13).type == ExcelError);
  
  
  // Test SUMIF
  // ... where there is only a check range
  assert(sumif_2(sumifs_array_1_v,new_excel_string(">0")).number == 110.0);
  assert(sumif_2(sumifs_array_1_v,new_excel_string(">10")).number == 100.0);
  assert(sumif_2(sumifs_array_1_v,new_excel_string("<100")).number == 10.0);
  
  // ... where there is a seprate sum range
  ExcelValue sumif_array_1[] = {new_excel_number(15),new_excel_number(20), new_excel_number(30)};
  ExcelValue sumif_array_1_v = new_excel_range(sumif_array_1,3,1);
  assert(sumif(sumifs_array_1_v,new_excel_string("10"),sumif_array_1_v).number == 15);
  
  
  // Test SUMPRODUCT
  ExcelValue sumproduct_1[] = { new_excel_number(10), new_excel_number(100), BLANK};
  ExcelValue sumproduct_2[] = { BLANK, new_excel_number(100), new_excel_number(10), BLANK};
  ExcelValue sumproduct_3[] = { BLANK };
  ExcelValue sumproduct_4[] = { new_excel_number(10), new_excel_number(100), new_excel_number(1000)};
  ExcelValue sumproduct_5[] = { new_excel_number(1), new_excel_number(2), new_excel_number(3)};
  ExcelValue sumproduct_6[] = { new_excel_number(1), new_excel_number(2), new_excel_number(4), new_excel_number(5)};
  ExcelValue sumproduct_7[] = { new_excel_number(10), new_excel_number(20), new_excel_number(40), new_excel_number(50)};
  ExcelValue sumproduct_8[] = { new_excel_number(11), new_excel_number(21), new_excel_number(41), new_excel_number(51)};
  ExcelValue sumproduct_9[] = { BLANK, BLANK };
  
  ExcelValue sumproduct_1_v = new_excel_range( sumproduct_1, 3, 1);
  ExcelValue sumproduct_2_v = new_excel_range( sumproduct_2, 3, 1);
  ExcelValue sumproduct_3_v = new_excel_range( sumproduct_3, 1, 1);
  // ExcelValue sumproduct_4_v = new_excel_range( sumproduct_4, 1, 3); // Unused
  ExcelValue sumproduct_5_v = new_excel_range( sumproduct_5, 3, 1);
  ExcelValue sumproduct_6_v = new_excel_range( sumproduct_6, 2, 2);
  ExcelValue sumproduct_7_v = new_excel_range( sumproduct_7, 2, 2);
  ExcelValue sumproduct_8_v = new_excel_range( sumproduct_8, 2, 2);
  ExcelValue sumproduct_9_v = new_excel_range( sumproduct_9, 2, 1);
  
  // ... should multiply together and then sum the elements in row or column areas given as arguments
  ExcelValue sumproducta_1[] = {sumproduct_1_v, sumproduct_2_v};
  assert(sumproduct(2,sumproducta_1).number == 100*100);

  // ... should return :value when miss-matched array sizes
  ExcelValue sumproducta_2[] = {sumproduct_1_v, sumproduct_3_v};
  assert(sumproduct(2,sumproducta_2).type == ExcelError);

  // ... if all its arguments are single values, should multiply them together
  // ExcelValue *sumproducta_3 = sumproduct_4;
  assert(sumproduct(3,sumproduct_4).number == 10*100*1000);

  // ... if it only has one range as an argument, should add its elements together
  ExcelValue sumproducta_4[] = {sumproduct_5_v};
  assert(sumproduct(1,sumproducta_4).number == 1 + 2 + 3);

  // ... if given multi row and column areas as arguments, should multipy the corresponding cell in each area and then add them all
  ExcelValue sumproducta_5[] = {sumproduct_6_v, sumproduct_7_v, sumproduct_8_v};
  assert(sumproduct(3,sumproducta_5).number == 1*10*11 + 2*20*21 + 4*40*41 + 5*50*51);

  // ... should raise an error if BLANK values outside of an array
  ExcelValue sumproducta_6[] = {BLANK,new_excel_number(1)};
  assert(sumproduct(2,sumproducta_6).type == ExcelError);

  // ... should ignore non-numeric values within an array
  ExcelValue sumproducta_7[] = {sumproduct_9_v, sumproduct_9_v};
  assert(sumproduct(2,sumproducta_7).number == 0);

  // ... should return an error if an argument is an error
  ExcelValue sumproducta_8[] = {VALUE};
  assert(sumproduct(1,sumproducta_8).type == ExcelError);
  
  // Test VLOOKUP
  ExcelValue vlookup_a1[] = {new_excel_number(1),new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a2[] = {new_excel_string("hello"),new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a3[] = {BLANK,new_excel_number(10),new_excel_number(2),new_excel_number(20),new_excel_number(3),new_excel_number(30)};
  ExcelValue vlookup_a1_v = new_excel_range(vlookup_a1,3,2);
  ExcelValue vlookup_a2_v = new_excel_range(vlookup_a2,3,2);
  ExcelValue vlookup_a3_v = new_excel_range(vlookup_a3,3,2);
  // ... should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument
  assert(vlookup_3(new_excel_number(2.0),vlookup_a1_v,new_excel_number(2)).number == 20);
  assert(vlookup_3(new_excel_number(1.5),vlookup_a1_v,new_excel_number(2)).number == 10);
  assert(vlookup_3(new_excel_number(0.5),vlookup_a1_v,new_excel_number(2)).type == ExcelError);
  assert(vlookup_3(new_excel_number(10),vlookup_a1_v,new_excel_number(2)).number == 30);
  assert(vlookup_3(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2)).number == 20);
  // ... has a four argument variant that matches the lookup type
  assert(vlookup(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2),TRUE).number == 20);
  assert(vlookup(new_excel_number(2.6),vlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_string("HELLO"),vlookup_a2_v,new_excel_number(2),FALSE).number == 10);
  assert(vlookup(new_excel_string("HELMP"),vlookup_a2_v,new_excel_number(2),TRUE).number == 10);
  // ... BLANK should not match with anything" do
  assert(vlookup_3(BLANK,vlookup_a3_v,new_excel_number(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(vlookup(VALUE,vlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),VALUE,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,new_excel_number(2),VALUE).type == ExcelError);
  assert(vlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);
	
  // Test SUM
  ExcelValue sum_array_0[] = {new_excel_number(1084.4557258064517),new_excel_number(32.0516914516129),new_excel_number(137.36439193548387)};
  ExcelValue sum_array_0_v = new_excel_range(sum_array_0,3,1);
  ExcelValue sum_array_1[] = {sum_array_0_v};
  assert(sum(1,sum_array_1).number == 1253.8718091935484);
  
  return 0;
}

int main() {
	return test_functions();
}
// End of the generic c functions

// Start of the file specific functions

// definitions
static ExcelValue _common0();
static ExcelValue _common1();
static ExcelValue _common2();
static ExcelValue _common3();
static ExcelValue _common4();
static ExcelValue _common5();
static ExcelValue _common6();
static ExcelValue _common7();
static ExcelValue _common8();
static ExcelValue _common9();
static ExcelValue _common10();
static ExcelValue _common11();
static ExcelValue _common12();
static ExcelValue _common13();
static ExcelValue _common14();
static ExcelValue _common15();
static ExcelValue _common16();
static ExcelValue _common17();
static ExcelValue _common18();
static ExcelValue _common19();
static ExcelValue _common20();
static ExcelValue _common21();
static ExcelValue _common22();
static ExcelValue _common23();
static ExcelValue _common24();
static ExcelValue _common25();
static ExcelValue _common26();
static ExcelValue _common27();
static ExcelValue _common28();
static ExcelValue _common29();
static ExcelValue _common30();
static ExcelValue _common31();
static ExcelValue _common32();
static ExcelValue _common33();
static ExcelValue _common34();
static ExcelValue _common35();
static ExcelValue _common36();
static ExcelValue _common37();
static ExcelValue _common38();
static ExcelValue _common39();
static ExcelValue _common40();
static ExcelValue _common41();
static ExcelValue _common42();
static ExcelValue _common43();
static ExcelValue _common44();
static ExcelValue _common45();
static ExcelValue _common46();
static ExcelValue _common47();
static ExcelValue _common48();
static ExcelValue _common49();
static ExcelValue _common50();
static ExcelValue _common51();
static ExcelValue _common52();
static ExcelValue _common53();
static ExcelValue _common54();
static ExcelValue _common55();
static ExcelValue _common56();
static ExcelValue _common57();
static ExcelValue _common58();
static ExcelValue _common59();
static ExcelValue _common60();
static ExcelValue _common61();
static ExcelValue _common62();
static ExcelValue _common63();
static ExcelValue _common64();
static ExcelValue _common65();
static ExcelValue _common66();
static ExcelValue _common67();
static ExcelValue _common68();
static ExcelValue _common69();
static ExcelValue _common70();
static ExcelValue _common71();
static ExcelValue _common72();
static ExcelValue _common73();
static ExcelValue _common74();
static ExcelValue _common75();
static ExcelValue _common76();
static ExcelValue _common77();
static ExcelValue _common78();
static ExcelValue _common79();
static ExcelValue _common80();
static ExcelValue _common81();
static ExcelValue _common82();
static ExcelValue _common83();
static ExcelValue _common84();
static ExcelValue _common85();
static ExcelValue _common86();
static ExcelValue _common87();
static ExcelValue _common88();
static ExcelValue _common89();
static ExcelValue _common90();
static ExcelValue _common91();
static ExcelValue _common92();
static ExcelValue _common93();
static ExcelValue _common94();
static ExcelValue _common95();
static ExcelValue _common96();
static ExcelValue _common97();
static ExcelValue _common98();
static ExcelValue _common99();
static ExcelValue _common100();
static ExcelValue _common101();
static ExcelValue _common102();
static ExcelValue _common103();
static ExcelValue _common104();
static ExcelValue _common105();
static ExcelValue _common106();
static ExcelValue _common107();
static ExcelValue _common108();
static ExcelValue _common109();
static ExcelValue _common110();
static ExcelValue _common111();
static ExcelValue _common112();
static ExcelValue _common113();
static ExcelValue _common114();
static ExcelValue _common115();
static ExcelValue _common116();
static ExcelValue _common117();
static ExcelValue _common118();
static ExcelValue _common119();
static ExcelValue _common120();
static ExcelValue _common121();
static ExcelValue _common122();
static ExcelValue _common123();
static ExcelValue _common124();
static ExcelValue _common125();
static ExcelValue _common126();
static ExcelValue _common127();
static ExcelValue _common128();
static ExcelValue _common129();
static ExcelValue _common130();
static ExcelValue _common131();
static ExcelValue _common132();
static ExcelValue _common133();
static ExcelValue _common134();
static ExcelValue _common135();
static ExcelValue _common136();
static ExcelValue _common137();
ExcelValue eu_b1();
ExcelValue eu_d2();
ExcelValue eu_d3();
ExcelValue eu_o3();
ExcelValue eu_b4();
ExcelValue eu_c4();
ExcelValue eu_d4();
ExcelValue eu_e4();
ExcelValue eu_f4();
ExcelValue eu_g4();
ExcelValue eu_h4();
ExcelValue eu_i4();
ExcelValue eu_j4();
ExcelValue eu_k4();
ExcelValue eu_l4();
ExcelValue eu_m4();
ExcelValue eu_o4();
ExcelValue eu_b5();
ExcelValue eu_c5();
ExcelValue eu_d5();
ExcelValue eu_e5();
ExcelValue eu_f5();
ExcelValue eu_g5();
ExcelValue eu_h5();
ExcelValue eu_i5();
ExcelValue eu_j5();
ExcelValue eu_k5();
ExcelValue eu_l5();
ExcelValue eu_m5();
ExcelValue eu_n5();
ExcelValue eu_o5();
ExcelValue eu_b6();
ExcelValue eu_c6();
ExcelValue eu_f6();
ExcelValue eu_g6();
ExcelValue eu_h6();
ExcelValue eu_i6();
ExcelValue eu_j6();
ExcelValue eu_k6();
ExcelValue eu_l6();
ExcelValue eu_m6();
ExcelValue eu_n6();
ExcelValue eu_o6();
ExcelValue eu_b7();
ExcelValue eu_c7();
ExcelValue eu_d7();
ExcelValue eu_e7();
ExcelValue eu_f7();
ExcelValue eu_g7();
ExcelValue eu_h7();
ExcelValue eu_i7();
ExcelValue eu_j7();
ExcelValue eu_k7();
ExcelValue eu_l7();
ExcelValue eu_m7();
ExcelValue eu_n7();
ExcelValue eu_o7();
ExcelValue eu_b8();
ExcelValue eu_f8();
ExcelValue eu_g8();
ExcelValue eu_h8();
ExcelValue eu_i8();
ExcelValue eu_j8();
ExcelValue eu_k8();
ExcelValue eu_l8();
ExcelValue eu_m8();
ExcelValue eu_n8();
ExcelValue eu_b10();
ExcelValue eu_f10();
ExcelValue eu_g10();
ExcelValue eu_h10();
ExcelValue eu_i10();
ExcelValue eu_j10();
ExcelValue eu_k10();
ExcelValue eu_l10();
ExcelValue eu_m10();
ExcelValue eu_b11();
ExcelValue eu_f11();
ExcelValue eu_g11();
ExcelValue eu_h11();
ExcelValue eu_i11();
ExcelValue eu_j11();
ExcelValue eu_k11();
ExcelValue eu_l11();
ExcelValue eu_m11();
ExcelValue eu_n11();
ExcelValue eu_b12();
ExcelValue eu_f12();
ExcelValue eu_g12();
ExcelValue eu_h12();
ExcelValue eu_i12();
ExcelValue eu_j12();
ExcelValue eu_k12();
ExcelValue eu_l12();
ExcelValue eu_m12();
ExcelValue eu_o12();
ExcelValue eu_b13();
ExcelValue eu_f13();
ExcelValue eu_g13();
ExcelValue eu_h13();
ExcelValue eu_i13();
ExcelValue eu_j13();
ExcelValue eu_k13();
ExcelValue eu_l13();
ExcelValue eu_m13();
ExcelValue eu_n13();
ExcelValue eu_o13();
ExcelValue eu_b14();
ExcelValue eu_f14();
ExcelValue eu_g14();
ExcelValue eu_h14();
ExcelValue eu_i14();
ExcelValue eu_j14();
ExcelValue eu_k14();
ExcelValue eu_l14();
ExcelValue eu_m14();
ExcelValue eu_b15();
ExcelValue eu_f15();
ExcelValue eu_g15();
ExcelValue eu_h15();
ExcelValue eu_i15();
ExcelValue eu_j15();
ExcelValue eu_k15();
ExcelValue eu_l15();
ExcelValue eu_m15();
ExcelValue eu_n15();
ExcelValue eu_b17();
ExcelValue eu_f17();
ExcelValue eu_g17();
ExcelValue eu_h17();
ExcelValue eu_i17();
ExcelValue eu_j17();
ExcelValue eu_k17();
ExcelValue eu_l17();
ExcelValue eu_m17();
ExcelValue eu_b18();
ExcelValue eu_d18();
ExcelValue eu_f18();
ExcelValue eu_g18();
ExcelValue eu_h18();
ExcelValue eu_i18();
ExcelValue eu_j18();
ExcelValue eu_k18();
ExcelValue eu_l18();
ExcelValue eu_m18();
ExcelValue eu_b19();
ExcelValue eu_d19();
ExcelValue eu_f19();
ExcelValue eu_g19();
ExcelValue eu_h19();
ExcelValue eu_i19();
ExcelValue eu_j19();
ExcelValue eu_k19();
ExcelValue eu_l19();
ExcelValue eu_m19();
ExcelValue eu_o19();
ExcelValue eu_p19();
ExcelValue eu_b20();
ExcelValue eu_f20();
ExcelValue eu_g20();
ExcelValue eu_h20();
ExcelValue eu_i20();
ExcelValue eu_j20();
ExcelValue eu_k20();
ExcelValue eu_l20();
ExcelValue eu_m20();
ExcelValue eu_b21();
ExcelValue eu_f21();
ExcelValue eu_g21();
ExcelValue eu_h21();
ExcelValue eu_i21();
ExcelValue eu_j21();
ExcelValue eu_k21();
ExcelValue eu_l21();
ExcelValue eu_m21();
ExcelValue eu_b23();
ExcelValue eu_f23();
ExcelValue eu_g23();
ExcelValue eu_h23();
ExcelValue eu_i23();
ExcelValue eu_j23();
ExcelValue eu_k23();
ExcelValue eu_l23();
ExcelValue eu_m23();
ExcelValue eu_b24();
ExcelValue eu_f24();
ExcelValue eu_g24();
ExcelValue eu_h24();
ExcelValue eu_i24();
ExcelValue eu_j24();
ExcelValue eu_k24();
ExcelValue eu_l24();
ExcelValue eu_m24();
ExcelValue eu_n24();
ExcelValue eu_o24();
ExcelValue eu_b25();
ExcelValue eu_f25();
ExcelValue eu_g25();
ExcelValue eu_h25();
ExcelValue eu_i25();
ExcelValue eu_j25();
ExcelValue eu_k25();
ExcelValue eu_l25();
ExcelValue eu_m25();
ExcelValue eu_b27();
ExcelValue eu_f27();
ExcelValue eu_g27();
ExcelValue eu_h27();
ExcelValue eu_i27();
ExcelValue eu_j27();
ExcelValue eu_k27();
ExcelValue eu_l27();
ExcelValue eu_m27();
ExcelValue eu_b28();
ExcelValue eu_f28();
ExcelValue eu_g28();
ExcelValue eu_h28();
ExcelValue eu_i28();
ExcelValue eu_j28();
ExcelValue eu_k28();
ExcelValue eu_l28();
ExcelValue eu_m28();
ExcelValue eu_b29();
ExcelValue eu_f29();
ExcelValue eu_g29();
ExcelValue eu_h29();
ExcelValue eu_i29();
ExcelValue eu_j29();
ExcelValue eu_k29();
ExcelValue eu_l29();
ExcelValue eu_m29();
ExcelValue eu_b30();
ExcelValue eu_f30();
ExcelValue eu_g30();
ExcelValue eu_h30();
ExcelValue eu_i30();
ExcelValue eu_j30();
ExcelValue eu_k30();
ExcelValue eu_l30();
ExcelValue eu_m30();
ExcelValue eu_b31();
ExcelValue eu_f31();
ExcelValue eu_g31();
ExcelValue eu_h31();
ExcelValue eu_i31();
ExcelValue eu_j31();
ExcelValue eu_k31();
ExcelValue eu_l31();
ExcelValue eu_m31();
ExcelValue eu_b33();
ExcelValue eu_b34();
ExcelValue eu_f34();
ExcelValue eu_g34();
ExcelValue eu_h34();
ExcelValue eu_i34();
ExcelValue eu_j34();
ExcelValue eu_k34();
ExcelValue eu_l34();
ExcelValue eu_m34();
ExcelValue eu_b35();
ExcelValue eu_f35();
ExcelValue eu_g35();
ExcelValue eu_h35();
ExcelValue eu_i35();
ExcelValue eu_j35();
ExcelValue eu_k35();
ExcelValue eu_l35();
ExcelValue eu_m35();
ExcelValue eu_b39();
ExcelValue eu_f39();
ExcelValue eu_g39();
ExcelValue eu_n39();
ExcelValue eu_b40();
ExcelValue eu_f40();
ExcelValue eu_g40();
ExcelValue eu_b41();
ExcelValue eu_f41();
ExcelValue eu_b42();
ExcelValue eu_f42();
ExcelValue eu_b43();
ExcelValue eu_f43();
ExcelValue eu_b45();
ExcelValue eu_f45();
ExcelValue eu_g45();
ExcelValue eu_h45();
ExcelValue eu_i45();
ExcelValue eu_j45();
ExcelValue eu_k45();
ExcelValue eu_l45();
ExcelValue eu_m45();
ExcelValue eu_b46();
ExcelValue eu_f46();
ExcelValue eu_g46();
ExcelValue eu_h46();
ExcelValue eu_i46();
ExcelValue eu_j46();
ExcelValue eu_k46();
ExcelValue eu_l46();
ExcelValue eu_m46();
ExcelValue eu_n46();
ExcelValue old_uk_b2();
ExcelValue old_uk_b4();
ExcelValue old_uk_c4();
ExcelValue old_uk_d4();
ExcelValue old_uk_e4();
ExcelValue old_uk_f4();
ExcelValue old_uk_g4();
ExcelValue old_uk_h4();
ExcelValue old_uk_i4();
ExcelValue old_uk_j4();
ExcelValue old_uk_l4();
ExcelValue old_uk_b5();
ExcelValue old_uk_c5();
ExcelValue old_uk_d5();
ExcelValue old_uk_e5();
ExcelValue old_uk_f5();
ExcelValue old_uk_g5();
ExcelValue old_uk_h5();
ExcelValue old_uk_i5();
ExcelValue old_uk_j5();
ExcelValue old_uk_k5();
ExcelValue old_uk_l5();
ExcelValue old_uk_b6();
ExcelValue old_uk_c6();
ExcelValue old_uk_d6();
ExcelValue old_uk_e6();
ExcelValue old_uk_f6();
ExcelValue old_uk_g6();
ExcelValue old_uk_h6();
ExcelValue old_uk_i6();
ExcelValue old_uk_j6();
ExcelValue old_uk_k6();
ExcelValue old_uk_l6();
ExcelValue old_uk_b7();
ExcelValue old_uk_c7();
ExcelValue old_uk_d7();
ExcelValue old_uk_e7();
ExcelValue old_uk_f7();
ExcelValue old_uk_g7();
ExcelValue old_uk_h7();
ExcelValue old_uk_i7();
ExcelValue old_uk_j7();
ExcelValue old_uk_k7();
ExcelValue old_uk_l7();
ExcelValue old_uk_b9();
ExcelValue old_uk_c9();
ExcelValue old_uk_d9();
ExcelValue old_uk_e9();
ExcelValue old_uk_f9();
ExcelValue old_uk_g9();
ExcelValue old_uk_h9();
ExcelValue old_uk_i9();
ExcelValue old_uk_j9();
ExcelValue old_uk_b10();
ExcelValue old_uk_c10();
ExcelValue old_uk_d10();
ExcelValue old_uk_e10();
ExcelValue old_uk_f10();
ExcelValue old_uk_g10();
ExcelValue old_uk_h10();
ExcelValue old_uk_i10();
ExcelValue old_uk_j10();
ExcelValue old_uk_k10();
ExcelValue old_uk_b11();
ExcelValue old_uk_c11();
ExcelValue old_uk_d11();
ExcelValue old_uk_e11();
ExcelValue old_uk_f11();
ExcelValue old_uk_g11();
ExcelValue old_uk_h11();
ExcelValue old_uk_i11();
ExcelValue old_uk_j11();
ExcelValue old_uk_k11();
ExcelValue old_uk_l11();
ExcelValue old_uk_b12();
ExcelValue old_uk_c12();
ExcelValue old_uk_d12();
ExcelValue old_uk_e12();
ExcelValue old_uk_f12();
ExcelValue old_uk_g12();
ExcelValue old_uk_h12();
ExcelValue old_uk_i12();
ExcelValue old_uk_j12();
ExcelValue old_uk_k12();
ExcelValue old_uk_c14();
ExcelValue old_uk_d14();
ExcelValue old_uk_e14();
ExcelValue old_uk_f14();
ExcelValue old_uk_g14();
ExcelValue old_uk_h14();
ExcelValue old_uk_i14();
ExcelValue old_uk_j14();
ExcelValue old_uk_b15();
ExcelValue old_uk_c15();
ExcelValue old_uk_d15();
ExcelValue old_uk_e15();
ExcelValue old_uk_f15();
ExcelValue old_uk_g15();
ExcelValue old_uk_h15();
ExcelValue old_uk_i15();
ExcelValue old_uk_j15();
ExcelValue old_uk_k15();
ExcelValue old_uk_b17();
ExcelValue old_uk_c17();
ExcelValue old_uk_d17();
ExcelValue old_uk_e17();
ExcelValue old_uk_f17();
ExcelValue old_uk_g17();
ExcelValue old_uk_h17();
ExcelValue old_uk_i17();
ExcelValue old_uk_j17();
ExcelValue old_uk_b18();
ExcelValue old_uk_c18();
ExcelValue old_uk_d18();
ExcelValue old_uk_e18();
ExcelValue old_uk_f18();
ExcelValue old_uk_g18();
ExcelValue old_uk_h18();
ExcelValue old_uk_i18();
ExcelValue old_uk_j18();
ExcelValue old_uk_k18();
ExcelValue old_uk_b19();
ExcelValue old_uk_c19();
ExcelValue old_uk_d19();
ExcelValue old_uk_e19();
ExcelValue old_uk_f19();
ExcelValue old_uk_g19();
ExcelValue old_uk_h19();
ExcelValue old_uk_i19();
ExcelValue old_uk_j19();
ExcelValue old_uk_k19();
ExcelValue old_uk_b20();
ExcelValue old_uk_c20();
ExcelValue old_uk_d20();
ExcelValue old_uk_e20();
ExcelValue old_uk_f20();
ExcelValue old_uk_g20();
ExcelValue old_uk_h20();
ExcelValue old_uk_i20();
ExcelValue old_uk_j20();
ExcelValue old_uk_k20();
ExcelValue old_uk_k22();
// end of definitions

// Used to decide whether to recalculate a cell
static int variable_set[580];
void reset() {
  int i;
  cell_counter = 0;
  for(i = 0; i < 580; i++) {
    variable_set[i] = 0;
  }
};

// starting the value constants
static ExcelValue C1 = {.type = ExcelString, .string = "Note, numbers not checked. Worry about power sector figure. Worry about the UK share of auction revenues."};
static ExcelValue C2 = {.type = ExcelString, .string = "Expansion"};
static ExcelValue C3 = {.type = ExcelString, .string = "Additions"};
static ExcelValue C4 = {.type = ExcelString, .string = "Annual"};
static ExcelValue C5 = {.type = ExcelString, .string = "EU-27 Emissions"};
static ExcelValue C6 = {.type = ExcelString, .string = "2005-6"};
static ExcelValue C7 = {.type = ExcelString, .string = "Phase II"};
static ExcelValue C8 = {.type = ExcelString, .string = "Phase III"};
static ExcelValue C9 = {.type = ExcelNumber, .number = 2013};
static ExcelValue C10 = {.type = ExcelNumber, .number = 2014};
static ExcelValue C11 = {.type = ExcelNumber, .number = 2015};
static ExcelValue C12 = {.type = ExcelNumber, .number = 2016};
static ExcelValue C13 = {.type = ExcelNumber, .number = 2017};
static ExcelValue C14 = {.type = ExcelNumber, .number = 2018};
static ExcelValue C15 = {.type = ExcelNumber, .number = 2019};
static ExcelValue C16 = {.type = ExcelNumber, .number = 2020};
static ExcelValue C17 = {.type = ExcelString, .string = "Change"};
static ExcelValue C18 = {.type = ExcelString, .string = "Power sector"};
static ExcelValue C19 = {.type = ExcelNumber, .number = 1150};
static ExcelValue C20 = {.type = ExcelNumber, .number = 50};
static ExcelValue C21 = {.type = ExcelNumber, .number = 1125.0};
static ExcelValue C22 = {.type = ExcelNumber, .number = 1102.5};
static ExcelValue C23 = {.type = ExcelNumber, .number = 1080.45};
static ExcelValue C24 = {.type = ExcelNumber, .number = 0.98};
static ExcelValue C25 = {.type = ExcelString, .string = "mtCO2"};
static ExcelValue C26 = {.type = ExcelNumber, .number = 0.02};
static ExcelValue C27 = {.type = ExcelString, .string = "Leakage sectors"};
static ExcelValue C28 = {.type = ExcelNumber, .number = 350};
static ExcelValue C29 = {.type = ExcelNumber, .number = 332.5};
static ExcelValue C30 = {.type = ExcelNumber, .number = 325.84999999999997};
static ExcelValue C31 = {.type = ExcelNumber, .number = 319.33299999999997};
static ExcelValue C32 = {.type = ExcelNumber, .number = 0.01};
static ExcelValue C33 = {.type = ExcelString, .string = "Other sectors"};
static ExcelValue C34 = {.type = ExcelNumber, .number = 550};
static ExcelValue C35 = {.type = ExcelNumber, .number = 100};
static ExcelValue C36 = {.type = ExcelNumber, .number = 712.5};
static ExcelValue C37 = {.type = ExcelNumber, .number = 698.25};
static ExcelValue C38 = {.type = ExcelNumber, .number = 684.285};
static ExcelValue C39 = {.type = ExcelString, .string = "Total"};
static ExcelValue C40 = {.type = ExcelNumber, .number = 2170.0};
static ExcelValue C41 = {.type = ExcelNumber, .number = 2126.6};
static ExcelValue C42 = {.type = ExcelString, .string = "Proportion allocated for free"};
static ExcelValue C43 = {.type = ExcelNumber, .number = 0};
static ExcelValue C44 = {.type = ExcelString, .string = "%"};
static ExcelValue C45 = {.type = ExcelNumber, .number = 0.9};
static ExcelValue C46 = {.type = ExcelNumber, .number = 0.875};
static ExcelValue C47 = {.type = ExcelNumber, .number = 0.85};
static ExcelValue C48 = {.type = ExcelNumber, .number = 0.825};
static ExcelValue C49 = {.type = ExcelNumber, .number = -0.025};
static ExcelValue C50 = {.type = ExcelString, .string = "Other"};
static ExcelValue C51 = {.type = ExcelNumber, .number = 0.8};
static ExcelValue C52 = {.type = ExcelNumber, .number = 0.6857142857142857};
static ExcelValue C53 = {.type = ExcelNumber, .number = 0.5714285714285714};
static ExcelValue C54 = {.type = ExcelNumber, .number = -0.1142857142857143};
static ExcelValue C55 = {.type = ExcelString, .string = "Average non-power sectors"};
static ExcelValue C56 = {.type = ExcelNumber, .number = 0.8318181818181818};
static ExcelValue C57 = {.type = ExcelNumber, .number = 0.7459415584415585};
static ExcelValue C58 = {.type = ExcelString, .string = "Total free allocation, % emissions"};
static ExcelValue C59 = {.type = ExcelNumber, .number = 0.40057603686635945};
static ExcelValue C60 = {.type = ExcelNumber, .number = 763.91875};
static ExcelValue C61 = {.type = ExcelString, .string = "Proportion auctioned, other sectors"};
static ExcelValue C62 = {.type = ExcelNumber, .number = 0.1681818181818182};
static ExcelValue C63 = {.type = ExcelNumber, .number = 1};
static ExcelValue C64 = {.type = ExcelString, .string = "Total allowances"};
static ExcelValue C65 = {.type = ExcelNumber, .number = 2010};
static ExcelValue C66 = {.type = ExcelNumber, .number = 2207.0};
static ExcelValue C67 = {.type = ExcelNumber, .number = 2091.7946};
static ExcelValue C68 = {.type = ExcelNumber, .number = 2055.39737396};
static ExcelValue C69 = {.type = ExcelNumber, .number = 2019.6334596530962};
static ExcelValue C70 = {.type = ExcelNumber, .number = 0.9826};
static ExcelValue C71 = {.type = ExcelNumber, .number = 0.0174};
static ExcelValue C72 = {.type = ExcelString, .string = "Note: 1720 \"based on current scope\""};
static ExcelValue C73 = {.type = ExcelString, .string = "Total free allocation, MtCO2"};
static ExcelValue C74 = {.type = ExcelNumber, .number = 848.7075677419355};
static ExcelValue C75 = {.type = ExcelString, .string = "Volume available for auctioning"};
static ExcelValue C76 = {.type = ExcelString, .string = "Carbon Price"};
static ExcelValue C77 = {.type = ExcelString, .string = "Price per allowance"};
static ExcelValue C78 = {.type = ExcelNumber, .number = 25};
static ExcelValue C79 = {.type = ExcelNumber, .number = 26.25};
static ExcelValue C80 = {.type = ExcelNumber, .number = 27.5625};
static ExcelValue C81 = {.type = ExcelNumber, .number = 28.940625};
static ExcelValue C82 = {.type = ExcelNumber, .number = 1.05};
static ExcelValue C83 = {.type = ExcelString, .string = "€/tCO2"};
static ExcelValue C84 = {.type = ExcelNumber, .number = 0.05};
static ExcelValue C85 = {.type = ExcelString, .string = "Total revenue from auctions"};
static ExcelValue C86 = {.type = ExcelString, .string = "EU-27 Auction volumes- bought into proportion to net shortfall"};
static ExcelValue C87 = {.type = ExcelNumber, .number = 1084.4557258064517};
static ExcelValue C88 = {.type = ExcelNumber, .number = 32.0516914516129};
static ExcelValue C89 = {.type = ExcelNumber, .number = 40.731249999999996};
static ExcelValue C90 = {.type = ExcelNumber, .number = 0.15000000000000002};
static ExcelValue C91 = {.type = ExcelNumber, .number = 137.36439193548387};
static ExcelValue C92 = {.type = ExcelNumber, .number = 219.45};
static ExcelValue C93 = {.type = ExcelString, .string = "Revenues"};
static ExcelValue C94 = {.type = ExcelString, .string = "2005 UK ETS Emissions"};
static ExcelValue C95 = {.type = ExcelNumber, .number = 242};
static ExcelValue C96 = {.type = ExcelString, .string = " "};
static ExcelValue C97 = {.type = ExcelString, .string = "2005 EU ETS Emissions"};
static ExcelValue C98 = {.type = ExcelNumber, .number = 1785};
static ExcelValue C99 = {.type = ExcelString, .string = "Basic UK share of auction revenues"};
static ExcelValue C100 = {.type = ExcelNumber, .number = 0.13557422969187674};
static ExcelValue C101 = {.type = ExcelString, .string = "Amount of share auctioned in UK"};
static ExcelValue C102 = {.type = ExcelString, .string = "Actual UK share of auction revenues"};
static ExcelValue C103 = {.type = ExcelNumber, .number = 0.12201680672268907};
static ExcelValue C104 = {.type = ExcelString, .string = "UK Auction revenues"};
static ExcelValue C105 = {.type = ExcelString, .string = "€bn"};
static ExcelValue C106 = {.type = ExcelString, .string = "UK - Possible phase III auction revenues"};
static ExcelValue C107 = {.type = ExcelString, .string = "Emissions"};
static ExcelValue C108 = {.type = ExcelNumber, .number = 163.25};
static ExcelValue C109 = {.type = ExcelNumber, .number = 160.40945};
static ExcelValue C110 = {.type = ExcelNumber, .number = 157.61832557};
static ExcelValue C111 = {.type = ExcelNumber, .number = 154.87576670508201};
static ExcelValue C112 = {.type = ExcelNumber, .number = 134.16697500000004};
static ExcelValue C113 = {.type = ExcelNumber, .number = 131.83246963500005};
static ExcelValue C114 = {.type = ExcelNumber, .number = 129.53858466335103};
static ExcelValue C115 = {.type = ExcelNumber, .number = 282.1603799952907};
static ExcelValue C116 = {.type = ExcelString, .string = "mtCO3"};
static ExcelValue C117 = {.type = ExcelNumber, .number = 0.14285714285714285};
static ExcelValue C118 = {.type = ExcelNumber, .number = 297.41697500000004};
static ExcelValue C119 = {.type = ExcelNumber, .number = 292.24191963500004};
static ExcelValue C120 = {.type = ExcelNumber, .number = 287.156910233351};
static ExcelValue C121 = {.type = ExcelString, .string = "mtCO4"};
static ExcelValue C122 = {.type = ExcelString, .string = "Proportion auctioned"};
static ExcelValue C123 = {.type = ExcelNumber, .number = 0.2};
static ExcelValue C124 = {.type = ExcelNumber, .number = 0.3142857142857143};
static ExcelValue C125 = {.type = ExcelNumber, .number = 0.4285714285714286};
static ExcelValue C126 = {.type = ExcelNumber, .number = 0.1142857142857143};
static ExcelValue C127 = {.type = ExcelNumber, .number = 0.6391141426947805};
static ExcelValue C128 = {.type = ExcelNumber, .number = 0.6906692651669547};
static ExcelValue C129 = {.type = ExcelString, .string = "CO2 price"};
static ExcelValue C130 = {.type = ExcelNumber, .number = 20};
static ExcelValue C131 = {.type = ExcelString, .string = "Auction revenues"};
static ExcelValue C132 = {.type = ExcelNumber, .number = 3.265};
static ExcelValue C133 = {.type = ExcelNumber, .number = 3.208189};
static ExcelValue C134 = {.type = ExcelNumber, .number = 3.1523665114};
static ExcelValue C135 = {.type = ExcelNumber, .number = 1000};
static ExcelValue C136 = {.type = ExcelNumber, .number = 0.5366679000000002};
static ExcelValue C137 = {.type = ExcelNumber, .number = 0.8286612377057146};
static ExcelValue C138 = {.type = ExcelNumber, .number = 3.8016679};
static ExcelValue C139 = {.type = ExcelNumber, .number = 290.9375};
static ExcelValue C140 = {.type = ExcelNumber, .number = 1004208.4312776};
static ExcelValue C141 = {.type = ExcelNumber, .number = 282.625};
static ExcelValue C142 = {.type = ExcelNumber, .number = 0.0};
// ending the value constants

// starting common elements
static ExcelValue _common0() {
  static ExcelValue result;
  if(variable_set[0] == 1) { return result;}
  result = divide(add(multiply(C31,C47),multiply(C38,C53)),add(C31,C38));
  variable_set[0] = 1;
  return result;
}

static ExcelValue _common1() {
  static ExcelValue result;
  if(variable_set[1] == 1) { return result;}
  result = add(multiply(C31,C47),multiply(C38,C53));
  variable_set[1] = 1;
  return result;
}

static ExcelValue _common2() {
  static ExcelValue result;
  if(variable_set[2] == 1) { return result;}
  result = add(multiply(C31,C47),multiply(C38,C53));
  variable_set[2] = 1;
  return result;
}

static ExcelValue _common3() {
  static ExcelValue result;
  if(variable_set[3] == 1) { return result;}
  result = multiply(C31,C47);
  variable_set[3] = 1;
  return result;
}

static ExcelValue _common4() {
  static ExcelValue result;
  if(variable_set[4] == 1) { return result;}
  result = multiply(C38,C53);
  variable_set[4] = 1;
  return result;
}

static ExcelValue _common5() {
  static ExcelValue result;
  if(variable_set[5] == 1) { return result;}
  result = add(C31,C38);
  variable_set[5] = 1;
  return result;
}

static ExcelValue _common6() {
  static ExcelValue result;
  if(variable_set[6] == 1) { return result;}
  result = divide(add(multiply(eu_i6(),C48),multiply(eu_i7(),eu_i13())),add(eu_i6(),eu_i7()));
  variable_set[6] = 1;
  return result;
}

static ExcelValue _common7() {
  static ExcelValue result;
  if(variable_set[7] == 1) { return result;}
  result = add(multiply(eu_i6(),C48),multiply(eu_i7(),eu_i13()));
  variable_set[7] = 1;
  return result;
}

static ExcelValue _common8() {
  static ExcelValue result;
  if(variable_set[8] == 1) { return result;}
  result = add(multiply(eu_i6(),C48),multiply(eu_i7(),eu_i13()));
  variable_set[8] = 1;
  return result;
}

static ExcelValue _common9() {
  static ExcelValue result;
  if(variable_set[9] == 1) { return result;}
  result = multiply(eu_i6(),C48);
  variable_set[9] = 1;
  return result;
}

static ExcelValue _common10() {
  static ExcelValue result;
  if(variable_set[10] == 1) { return result;}
  result = multiply(eu_i7(),eu_i13());
  variable_set[10] = 1;
  return result;
}

static ExcelValue _common11() {
  static ExcelValue result;
  if(variable_set[11] == 1) { return result;}
  result = add(eu_i6(),eu_i7());
  variable_set[11] = 1;
  return result;
}

static ExcelValue _common12() {
  static ExcelValue result;
  if(variable_set[12] == 1) { return result;}
  result = divide(add(multiply(eu_j6(),eu_j12()),multiply(eu_j7(),eu_j13())),add(eu_j6(),eu_j7()));
  variable_set[12] = 1;
  return result;
}

static ExcelValue _common13() {
  static ExcelValue result;
  if(variable_set[13] == 1) { return result;}
  result = add(multiply(eu_j6(),eu_j12()),multiply(eu_j7(),eu_j13()));
  variable_set[13] = 1;
  return result;
}

static ExcelValue _common14() {
  static ExcelValue result;
  if(variable_set[14] == 1) { return result;}
  result = add(multiply(eu_j6(),eu_j12()),multiply(eu_j7(),eu_j13()));
  variable_set[14] = 1;
  return result;
}

static ExcelValue _common15() {
  static ExcelValue result;
  if(variable_set[15] == 1) { return result;}
  result = multiply(eu_j6(),eu_j12());
  variable_set[15] = 1;
  return result;
}

static ExcelValue _common16() {
  static ExcelValue result;
  if(variable_set[16] == 1) { return result;}
  result = multiply(eu_j7(),eu_j13());
  variable_set[16] = 1;
  return result;
}

static ExcelValue _common17() {
  static ExcelValue result;
  if(variable_set[17] == 1) { return result;}
  result = add(eu_j6(),eu_j7());
  variable_set[17] = 1;
  return result;
}

static ExcelValue _common18() {
  static ExcelValue result;
  if(variable_set[18] == 1) { return result;}
  result = divide(add(multiply(eu_k6(),eu_k12()),multiply(eu_k7(),eu_k13())),add(eu_k6(),eu_k7()));
  variable_set[18] = 1;
  return result;
}

static ExcelValue _common19() {
  static ExcelValue result;
  if(variable_set[19] == 1) { return result;}
  result = add(multiply(eu_k6(),eu_k12()),multiply(eu_k7(),eu_k13()));
  variable_set[19] = 1;
  return result;
}

static ExcelValue _common20() {
  static ExcelValue result;
  if(variable_set[20] == 1) { return result;}
  result = add(multiply(eu_k6(),eu_k12()),multiply(eu_k7(),eu_k13()));
  variable_set[20] = 1;
  return result;
}

static ExcelValue _common21() {
  static ExcelValue result;
  if(variable_set[21] == 1) { return result;}
  result = multiply(eu_k6(),eu_k12());
  variable_set[21] = 1;
  return result;
}

static ExcelValue _common22() {
  static ExcelValue result;
  if(variable_set[22] == 1) { return result;}
  result = multiply(eu_k7(),eu_k13());
  variable_set[22] = 1;
  return result;
}

static ExcelValue _common23() {
  static ExcelValue result;
  if(variable_set[23] == 1) { return result;}
  result = add(eu_k6(),eu_k7());
  variable_set[23] = 1;
  return result;
}

static ExcelValue _common24() {
  static ExcelValue result;
  if(variable_set[24] == 1) { return result;}
  result = divide(add(multiply(eu_l6(),eu_l12()),multiply(eu_l7(),eu_l13())),add(eu_l6(),eu_l7()));
  variable_set[24] = 1;
  return result;
}

static ExcelValue _common25() {
  static ExcelValue result;
  if(variable_set[25] == 1) { return result;}
  result = add(multiply(eu_l6(),eu_l12()),multiply(eu_l7(),eu_l13()));
  variable_set[25] = 1;
  return result;
}

static ExcelValue _common26() {
  static ExcelValue result;
  if(variable_set[26] == 1) { return result;}
  result = add(multiply(eu_l6(),eu_l12()),multiply(eu_l7(),eu_l13()));
  variable_set[26] = 1;
  return result;
}

static ExcelValue _common27() {
  static ExcelValue result;
  if(variable_set[27] == 1) { return result;}
  result = multiply(eu_l6(),eu_l12());
  variable_set[27] = 1;
  return result;
}

static ExcelValue _common28() {
  static ExcelValue result;
  if(variable_set[28] == 1) { return result;}
  result = multiply(eu_l7(),eu_l13());
  variable_set[28] = 1;
  return result;
}

static ExcelValue _common29() {
  static ExcelValue result;
  if(variable_set[29] == 1) { return result;}
  result = add(eu_l6(),eu_l7());
  variable_set[29] = 1;
  return result;
}

static ExcelValue _common30() {
  static ExcelValue result;
  if(variable_set[30] == 1) { return result;}
  result = divide(add(multiply(eu_m6(),eu_m12()),multiply(eu_m7(),C43)),add(eu_m6(),eu_m7()));
  variable_set[30] = 1;
  return result;
}

static ExcelValue _common31() {
  static ExcelValue result;
  if(variable_set[31] == 1) { return result;}
  result = add(multiply(eu_m6(),eu_m12()),multiply(eu_m7(),C43));
  variable_set[31] = 1;
  return result;
}

static ExcelValue _common32() {
  static ExcelValue result;
  if(variable_set[32] == 1) { return result;}
  result = add(multiply(eu_m6(),eu_m12()),multiply(eu_m7(),C43));
  variable_set[32] = 1;
  return result;
}

static ExcelValue _common33() {
  static ExcelValue result;
  if(variable_set[33] == 1) { return result;}
  result = multiply(eu_m6(),eu_m12());
  variable_set[33] = 1;
  return result;
}

static ExcelValue _common34() {
  static ExcelValue result;
  if(variable_set[34] == 1) { return result;}
  result = multiply(eu_m7(),C43);
  variable_set[34] = 1;
  return result;
}

static ExcelValue _common35() {
  static ExcelValue result;
  if(variable_set[35] == 1) { return result;}
  result = add(eu_m6(),eu_m7());
  variable_set[35] = 1;
  return result;
}

static ExcelValue _common36() {
  static ExcelValue result;
  if(variable_set[36] == 1) { return result;}
  result = C69;
  variable_set[36] = 1;
  return result;
}

static ExcelValue _common37() {
  static ExcelValue result;
  if(variable_set[37] == 1) { return result;}
  result = eu_i19();
  variable_set[37] = 1;
  return result;
}

static ExcelValue _common38() {
  static ExcelValue result;
  if(variable_set[38] == 1) { return result;}
  result = eu_j19();
  variable_set[38] = 1;
  return result;
}

static ExcelValue _common39() {
  static ExcelValue result;
  if(variable_set[39] == 1) { return result;}
  result = eu_k19();
  variable_set[39] = 1;
  return result;
}

static ExcelValue _common40() {
  static ExcelValue result;
  if(variable_set[40] == 1) { return result;}
  result = eu_l19();
  variable_set[40] = 1;
  return result;
}

static ExcelValue _common41() {
  static ExcelValue result;
  if(variable_set[41] == 1) { return result;}
  result = add(C139,divide(C140,C41));
  variable_set[41] = 1;
  return result;
}

static ExcelValue _common42() {
  static ExcelValue result;
  if(variable_set[42] == 1) { return result;}
  result = divide(C140,C41);
  variable_set[42] = 1;
  return result;
}

static ExcelValue _common43() {
  static ExcelValue result;
  if(variable_set[43] == 1) { return result;}
  result = add(C141,divide(multiply(multiply(C36,C53),C69),eu_h8()));
  variable_set[43] = 1;
  return result;
}

static ExcelValue _common44() {
  static ExcelValue result;
  if(variable_set[44] == 1) { return result;}
  result = divide(multiply(multiply(C36,C53),C69),eu_h8());
  variable_set[44] = 1;
  return result;
}

static ExcelValue _common45() {
  static ExcelValue result;
  if(variable_set[45] == 1) { return result;}
  result = multiply(multiply(C36,C53),C69);
  variable_set[45] = 1;
  return result;
}

static ExcelValue _common46() {
  static ExcelValue result;
  if(variable_set[46] == 1) { return result;}
  result = multiply(C36,C53);
  variable_set[46] = 1;
  return result;
}

static ExcelValue _common47() {
  static ExcelValue result;
  if(variable_set[47] == 1) { return result;}
  result = add(multiply(C29,C48),divide(multiply(multiply(C36,eu_i13()),eu_i19()),eu_i8()));
  variable_set[47] = 1;
  return result;
}

static ExcelValue _common48() {
  static ExcelValue result;
  if(variable_set[48] == 1) { return result;}
  result = multiply(C29,C48);
  variable_set[48] = 1;
  return result;
}

static ExcelValue _common49() {
  static ExcelValue result;
  if(variable_set[49] == 1) { return result;}
  result = divide(multiply(multiply(C36,eu_i13()),eu_i19()),eu_i8());
  variable_set[49] = 1;
  return result;
}

static ExcelValue _common50() {
  static ExcelValue result;
  if(variable_set[50] == 1) { return result;}
  result = multiply(multiply(C36,eu_i13()),eu_i19());
  variable_set[50] = 1;
  return result;
}

static ExcelValue _common51() {
  static ExcelValue result;
  if(variable_set[51] == 1) { return result;}
  result = multiply(C36,eu_i13());
  variable_set[51] = 1;
  return result;
}

static ExcelValue _common52() {
  static ExcelValue result;
  if(variable_set[52] == 1) { return result;}
  result = add(multiply(C29,eu_j12()),divide(multiply(multiply(C36,eu_j13()),eu_j19()),eu_j8()));
  variable_set[52] = 1;
  return result;
}

static ExcelValue _common53() {
  static ExcelValue result;
  if(variable_set[53] == 1) { return result;}
  result = multiply(C29,eu_j12());
  variable_set[53] = 1;
  return result;
}

static ExcelValue _common54() {
  static ExcelValue result;
  if(variable_set[54] == 1) { return result;}
  result = divide(multiply(multiply(C36,eu_j13()),eu_j19()),eu_j8());
  variable_set[54] = 1;
  return result;
}

static ExcelValue _common55() {
  static ExcelValue result;
  if(variable_set[55] == 1) { return result;}
  result = multiply(multiply(C36,eu_j13()),eu_j19());
  variable_set[55] = 1;
  return result;
}

static ExcelValue _common56() {
  static ExcelValue result;
  if(variable_set[56] == 1) { return result;}
  result = multiply(C36,eu_j13());
  variable_set[56] = 1;
  return result;
}

static ExcelValue _common57() {
  static ExcelValue result;
  if(variable_set[57] == 1) { return result;}
  result = add(multiply(C29,eu_k12()),divide(multiply(multiply(C36,eu_k13()),eu_k19()),eu_k8()));
  variable_set[57] = 1;
  return result;
}

static ExcelValue _common58() {
  static ExcelValue result;
  if(variable_set[58] == 1) { return result;}
  result = multiply(C29,eu_k12());
  variable_set[58] = 1;
  return result;
}

static ExcelValue _common59() {
  static ExcelValue result;
  if(variable_set[59] == 1) { return result;}
  result = divide(multiply(multiply(C36,eu_k13()),eu_k19()),eu_k8());
  variable_set[59] = 1;
  return result;
}

static ExcelValue _common60() {
  static ExcelValue result;
  if(variable_set[60] == 1) { return result;}
  result = multiply(multiply(C36,eu_k13()),eu_k19());
  variable_set[60] = 1;
  return result;
}

static ExcelValue _common61() {
  static ExcelValue result;
  if(variable_set[61] == 1) { return result;}
  result = multiply(C36,eu_k13());
  variable_set[61] = 1;
  return result;
}

static ExcelValue _common62() {
  static ExcelValue result;
  if(variable_set[62] == 1) { return result;}
  result = add(multiply(C29,eu_l12()),divide(multiply(multiply(C36,eu_l13()),eu_l19()),eu_l8()));
  variable_set[62] = 1;
  return result;
}

static ExcelValue _common63() {
  static ExcelValue result;
  if(variable_set[63] == 1) { return result;}
  result = multiply(C29,eu_l12());
  variable_set[63] = 1;
  return result;
}

static ExcelValue _common64() {
  static ExcelValue result;
  if(variable_set[64] == 1) { return result;}
  result = divide(multiply(multiply(C36,eu_l13()),eu_l19()),eu_l8());
  variable_set[64] = 1;
  return result;
}

static ExcelValue _common65() {
  static ExcelValue result;
  if(variable_set[65] == 1) { return result;}
  result = multiply(multiply(C36,eu_l13()),eu_l19());
  variable_set[65] = 1;
  return result;
}

static ExcelValue _common66() {
  static ExcelValue result;
  if(variable_set[66] == 1) { return result;}
  result = multiply(C36,eu_l13());
  variable_set[66] = 1;
  return result;
}

static ExcelValue _common67() {
  static ExcelValue result;
  if(variable_set[67] == 1) { return result;}
  result = add(multiply(C29,eu_m12()),divide(multiply(C142,eu_m19()),eu_m8()));
  variable_set[67] = 1;
  return result;
}

static ExcelValue _common68() {
  static ExcelValue result;
  if(variable_set[68] == 1) { return result;}
  result = multiply(C29,eu_m12());
  variable_set[68] = 1;
  return result;
}

static ExcelValue _common69() {
  static ExcelValue result;
  if(variable_set[69] == 1) { return result;}
  result = divide(multiply(C142,eu_m19()),eu_m8());
  variable_set[69] = 1;
  return result;
}

static ExcelValue _common70() {
  static ExcelValue result;
  if(variable_set[70] == 1) { return result;}
  result = multiply(C142,eu_m19());
  variable_set[70] = 1;
  return result;
}

static ExcelValue _common71() {
  static ExcelValue result;
  if(variable_set[71] == 1) { return result;}
  result = subtract(C67,C74);
  variable_set[71] = 1;
  return result;
}

static ExcelValue _common72() {
  static ExcelValue result;
  if(variable_set[72] == 1) { return result;}
  result = subtract(C68,add(C139,divide(C140,C41)));
  variable_set[72] = 1;
  return result;
}

static ExcelValue _common73() {
  static ExcelValue result;
  if(variable_set[73] == 1) { return result;}
  result = subtract(C69,add(C141,divide(multiply(multiply(C36,C53),C69),eu_h8())));
  variable_set[73] = 1;
  return result;
}

static ExcelValue _common74() {
  static ExcelValue result;
  if(variable_set[74] == 1) { return result;}
  result = subtract(eu_i19(),add(multiply(C29,C48),divide(multiply(multiply(C36,eu_i13()),eu_i19()),eu_i8())));
  variable_set[74] = 1;
  return result;
}

static ExcelValue _common75() {
  static ExcelValue result;
  if(variable_set[75] == 1) { return result;}
  result = subtract(eu_j19(),add(multiply(C29,eu_j12()),divide(multiply(multiply(C36,eu_j13()),eu_j19()),eu_j8())));
  variable_set[75] = 1;
  return result;
}

static ExcelValue _common76() {
  static ExcelValue result;
  if(variable_set[76] == 1) { return result;}
  result = subtract(eu_k19(),add(multiply(C29,eu_k12()),divide(multiply(multiply(C36,eu_k13()),eu_k19()),eu_k8())));
  variable_set[76] = 1;
  return result;
}

static ExcelValue _common77() {
  static ExcelValue result;
  if(variable_set[77] == 1) { return result;}
  result = subtract(eu_l19(),add(multiply(C29,eu_l12()),divide(multiply(multiply(C36,eu_l13()),eu_l19()),eu_l8())));
  variable_set[77] = 1;
  return result;
}

static ExcelValue _common78() {
  static ExcelValue result;
  if(variable_set[78] == 1) { return result;}
  result = subtract(eu_m19(),add(multiply(C29,eu_m12()),divide(multiply(C142,eu_m19()),eu_m8())));
  variable_set[78] = 1;
  return result;
}

static ExcelValue _common79() {
  static ExcelValue result;
  if(variable_set[79] == 1) { return result;}
  result = eu_m19();
  variable_set[79] = 1;
  return result;
}

static ExcelValue _common80() {
  static ExcelValue result;
  if(variable_set[80] == 1) { return result;}
  result = C81;
  variable_set[80] = 1;
  return result;
}

static ExcelValue _common81() {
  static ExcelValue result;
  if(variable_set[81] == 1) { return result;}
  result = eu_j24();
  variable_set[81] = 1;
  return result;
}

static ExcelValue _common82() {
  static ExcelValue result;
  if(variable_set[82] == 1) { return result;}
  result = eu_k24();
  variable_set[82] = 1;
  return result;
}

static ExcelValue _common83() {
  static ExcelValue result;
  if(variable_set[83] == 1) { return result;}
  result = eu_l24();
  variable_set[83] = 1;
  return result;
}

static ExcelValue _common84() {
  static ExcelValue result;
  if(variable_set[84] == 1) { return result;}
  result = divide(C68,C41);
  variable_set[84] = 1;
  return result;
}

static ExcelValue _common85() {
  static ExcelValue result;
  if(variable_set[85] == 1) { return result;}
  result = divide(C69,eu_h8());
  variable_set[85] = 1;
  return result;
}

static ExcelValue _common86() {
  static ExcelValue result;
  if(variable_set[86] == 1) { return result;}
  result = divide(eu_i19(),eu_i8());
  variable_set[86] = 1;
  return result;
}

static ExcelValue _common87() {
  static ExcelValue result;
  if(variable_set[87] == 1) { return result;}
  result = divide(eu_j19(),eu_j8());
  variable_set[87] = 1;
  return result;
}

static ExcelValue _common88() {
  static ExcelValue result;
  if(variable_set[88] == 1) { return result;}
  result = divide(eu_k19(),eu_k8());
  variable_set[88] = 1;
  return result;
}

static ExcelValue _common89() {
  static ExcelValue result;
  if(variable_set[89] == 1) { return result;}
  result = divide(eu_l19(),eu_l8());
  variable_set[89] = 1;
  return result;
}

static ExcelValue _common90() {
  static ExcelValue result;
  if(variable_set[90] == 1) { return result;}
  result = divide(eu_m19(),eu_m8());
  variable_set[90] = 1;
  return result;
}

static ExcelValue _common91() {
  static ExcelValue result;
  if(variable_set[91] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = C87;
  array1[1] = C88;
  array1[2] = C91;
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[91] = 1;
  return result;
}

static ExcelValue _common92() {
  static ExcelValue result;
  if(variable_set[92] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = C87;
  array1[1] = C88;
  array1[2] = C91;
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[92] = 1;
  return result;
}

static ExcelValue _common93() {
  static ExcelValue result;
  if(variable_set[93] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = C87;
  array0[1] = C88;
  array0[2] = C91;
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[93] = 1;
  return result;
}

static ExcelValue _common94() {
  static ExcelValue result;
  if(variable_set[94] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_g28();
  array1[1] = eu_g29();
  array1[2] = eu_g30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[94] = 1;
  return result;
}

static ExcelValue _common95() {
  static ExcelValue result;
  if(variable_set[95] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_g28();
  array1[1] = eu_g29();
  array1[2] = eu_g30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[95] = 1;
  return result;
}

static ExcelValue _common96() {
  static ExcelValue result;
  if(variable_set[96] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_g28();
  array0[1] = eu_g29();
  array0[2] = eu_g30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[96] = 1;
  return result;
}

static ExcelValue _common97() {
  static ExcelValue result;
  if(variable_set[97] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_h28();
  array1[1] = eu_h29();
  array1[2] = eu_h30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[97] = 1;
  return result;
}

static ExcelValue _common98() {
  static ExcelValue result;
  if(variable_set[98] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_h28();
  array1[1] = eu_h29();
  array1[2] = eu_h30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[98] = 1;
  return result;
}

static ExcelValue _common99() {
  static ExcelValue result;
  if(variable_set[99] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_h28();
  array0[1] = eu_h29();
  array0[2] = eu_h30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[99] = 1;
  return result;
}

static ExcelValue _common100() {
  static ExcelValue result;
  if(variable_set[100] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_i28();
  array1[1] = eu_i29();
  array1[2] = eu_i30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[100] = 1;
  return result;
}

static ExcelValue _common101() {
  static ExcelValue result;
  if(variable_set[101] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_i28();
  array1[1] = eu_i29();
  array1[2] = eu_i30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[101] = 1;
  return result;
}

static ExcelValue _common102() {
  static ExcelValue result;
  if(variable_set[102] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_i28();
  array0[1] = eu_i29();
  array0[2] = eu_i30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[102] = 1;
  return result;
}

static ExcelValue _common103() {
  static ExcelValue result;
  if(variable_set[103] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_j28();
  array1[1] = eu_j29();
  array1[2] = eu_j30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[103] = 1;
  return result;
}

static ExcelValue _common104() {
  static ExcelValue result;
  if(variable_set[104] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_j28();
  array1[1] = eu_j29();
  array1[2] = eu_j30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[104] = 1;
  return result;
}

static ExcelValue _common105() {
  static ExcelValue result;
  if(variable_set[105] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_j28();
  array0[1] = eu_j29();
  array0[2] = eu_j30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[105] = 1;
  return result;
}

static ExcelValue _common106() {
  static ExcelValue result;
  if(variable_set[106] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_k28();
  array1[1] = eu_k29();
  array1[2] = eu_k30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[106] = 1;
  return result;
}

static ExcelValue _common107() {
  static ExcelValue result;
  if(variable_set[107] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_k28();
  array1[1] = eu_k29();
  array1[2] = eu_k30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[107] = 1;
  return result;
}

static ExcelValue _common108() {
  static ExcelValue result;
  if(variable_set[108] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_k28();
  array0[1] = eu_k29();
  array0[2] = eu_k30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[108] = 1;
  return result;
}

static ExcelValue _common109() {
  static ExcelValue result;
  if(variable_set[109] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_l28();
  array1[1] = eu_l29();
  array1[2] = eu_l30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[109] = 1;
  return result;
}

static ExcelValue _common110() {
  static ExcelValue result;
  if(variable_set[110] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_l28();
  array1[1] = eu_l29();
  array1[2] = eu_l30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[110] = 1;
  return result;
}

static ExcelValue _common111() {
  static ExcelValue result;
  if(variable_set[111] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_l28();
  array0[1] = eu_l29();
  array0[2] = eu_l30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[111] = 1;
  return result;
}

static ExcelValue _common112() {
  static ExcelValue result;
  if(variable_set[112] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_m28();
  array1[1] = eu_m29();
  array1[2] = eu_m30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[112] = 1;
  return result;
}

static ExcelValue _common113() {
  static ExcelValue result;
  if(variable_set[113] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_m28();
  array1[1] = eu_m29();
  array1[2] = eu_m30();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[113] = 1;
  return result;
}

static ExcelValue _common114() {
  static ExcelValue result;
  if(variable_set[114] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = eu_m28();
  array0[1] = eu_m29();
  array0[2] = eu_m30();
  ExcelValue array0_ev = new_excel_range(array0,3,1);
  result = array0_ev;
  variable_set[114] = 1;
  return result;
}

static ExcelValue _common115() {
  static ExcelValue result;
  if(variable_set[115] == 1) { return result;}
  result = divide(add(C110,multiply(C125,C114)),C120);
  variable_set[115] = 1;
  return result;
}

static ExcelValue _common116() {
  static ExcelValue result;
  if(variable_set[116] == 1) { return result;}
  result = add(C110,multiply(C125,C114));
  variable_set[116] = 1;
  return result;
}

static ExcelValue _common117() {
  static ExcelValue result;
  if(variable_set[117] == 1) { return result;}
  result = multiply(C125,C114);
  variable_set[117] = 1;
  return result;
}

static ExcelValue _common118() {
  static ExcelValue result;
  if(variable_set[118] == 1) { return result;}
  result = divide(add(multiply(C63,C111),multiply(old_uk_f11(),old_uk_f6())),C115);
  variable_set[118] = 1;
  return result;
}

static ExcelValue _common119() {
  static ExcelValue result;
  if(variable_set[119] == 1) { return result;}
  result = add(multiply(C63,C111),multiply(old_uk_f11(),old_uk_f6()));
  variable_set[119] = 1;
  return result;
}

static ExcelValue _common120() {
  static ExcelValue result;
  if(variable_set[120] == 1) { return result;}
  result = multiply(C63,C111);
  variable_set[120] = 1;
  return result;
}

static ExcelValue _common121() {
  static ExcelValue result;
  if(variable_set[121] == 1) { return result;}
  result = multiply(old_uk_f11(),old_uk_f6());
  variable_set[121] = 1;
  return result;
}

static ExcelValue _common122() {
  static ExcelValue result;
  if(variable_set[122] == 1) { return result;}
  result = divide(add(multiply(C63,old_uk_g5()),multiply(old_uk_g11(),old_uk_g6())),old_uk_g7());
  variable_set[122] = 1;
  return result;
}

static ExcelValue _common123() {
  static ExcelValue result;
  if(variable_set[123] == 1) { return result;}
  result = add(multiply(C63,old_uk_g5()),multiply(old_uk_g11(),old_uk_g6()));
  variable_set[123] = 1;
  return result;
}

static ExcelValue _common124() {
  static ExcelValue result;
  if(variable_set[124] == 1) { return result;}
  result = multiply(C63,old_uk_g5());
  variable_set[124] = 1;
  return result;
}

static ExcelValue _common125() {
  static ExcelValue result;
  if(variable_set[125] == 1) { return result;}
  result = multiply(old_uk_g11(),old_uk_g6());
  variable_set[125] = 1;
  return result;
}

static ExcelValue _common126() {
  static ExcelValue result;
  if(variable_set[126] == 1) { return result;}
  result = divide(add(multiply(C63,old_uk_h5()),multiply(old_uk_h11(),old_uk_h6())),old_uk_h7());
  variable_set[126] = 1;
  return result;
}

static ExcelValue _common127() {
  static ExcelValue result;
  if(variable_set[127] == 1) { return result;}
  result = add(multiply(C63,old_uk_h5()),multiply(old_uk_h11(),old_uk_h6()));
  variable_set[127] = 1;
  return result;
}

static ExcelValue _common128() {
  static ExcelValue result;
  if(variable_set[128] == 1) { return result;}
  result = multiply(C63,old_uk_h5());
  variable_set[128] = 1;
  return result;
}

static ExcelValue _common129() {
  static ExcelValue result;
  if(variable_set[129] == 1) { return result;}
  result = multiply(old_uk_h11(),old_uk_h6());
  variable_set[129] = 1;
  return result;
}

static ExcelValue _common130() {
  static ExcelValue result;
  if(variable_set[130] == 1) { return result;}
  result = divide(add(multiply(C63,old_uk_i5()),multiply(old_uk_i11(),old_uk_i6())),old_uk_i7());
  variable_set[130] = 1;
  return result;
}

static ExcelValue _common131() {
  static ExcelValue result;
  if(variable_set[131] == 1) { return result;}
  result = add(multiply(C63,old_uk_i5()),multiply(old_uk_i11(),old_uk_i6()));
  variable_set[131] = 1;
  return result;
}

static ExcelValue _common132() {
  static ExcelValue result;
  if(variable_set[132] == 1) { return result;}
  result = multiply(C63,old_uk_i5());
  variable_set[132] = 1;
  return result;
}

static ExcelValue _common133() {
  static ExcelValue result;
  if(variable_set[133] == 1) { return result;}
  result = multiply(old_uk_i11(),old_uk_i6());
  variable_set[133] = 1;
  return result;
}

static ExcelValue _common134() {
  static ExcelValue result;
  if(variable_set[134] == 1) { return result;}
  result = divide(add(multiply(C63,old_uk_j5()),multiply(C63,old_uk_j6())),old_uk_j7());
  variable_set[134] = 1;
  return result;
}

static ExcelValue _common135() {
  static ExcelValue result;
  if(variable_set[135] == 1) { return result;}
  result = add(multiply(C63,old_uk_j5()),multiply(C63,old_uk_j6()));
  variable_set[135] = 1;
  return result;
}

static ExcelValue _common136() {
  static ExcelValue result;
  if(variable_set[136] == 1) { return result;}
  result = multiply(C63,old_uk_j5());
  variable_set[136] = 1;
  return result;
}

static ExcelValue _common137() {
  static ExcelValue result;
  if(variable_set[137] == 1) { return result;}
  result = multiply(C63,old_uk_j6());
  variable_set[137] = 1;
  return result;
}

// ending common elements

// start EU
ExcelValue eu_b1() {
  static ExcelValue result;
  if(variable_set[138] == 1) { return result;}
  result = C1;
  variable_set[138] = 1;
  return result;
}

ExcelValue eu_d2() {
  static ExcelValue result;
  if(variable_set[139] == 1) { return result;}
  result = C2;
  variable_set[139] = 1;
  return result;
}

ExcelValue eu_d3() {
  static ExcelValue result;
  if(variable_set[140] == 1) { return result;}
  result = C3;
  variable_set[140] = 1;
  return result;
}

ExcelValue eu_o3() {
  static ExcelValue result;
  if(variable_set[141] == 1) { return result;}
  result = C4;
  variable_set[141] = 1;
  return result;
}

ExcelValue eu_b4() {
  static ExcelValue result;
  if(variable_set[142] == 1) { return result;}
  result = C5;
  variable_set[142] = 1;
  return result;
}

ExcelValue eu_c4() {
  static ExcelValue result;
  if(variable_set[143] == 1) { return result;}
  result = C6;
  variable_set[143] = 1;
  return result;
}

ExcelValue eu_d4() {
  static ExcelValue result;
  if(variable_set[144] == 1) { return result;}
  result = C7;
  variable_set[144] = 1;
  return result;
}

ExcelValue eu_e4() {
  static ExcelValue result;
  if(variable_set[145] == 1) { return result;}
  result = C8;
  variable_set[145] = 1;
  return result;
}

ExcelValue eu_f4() {
  static ExcelValue result;
  if(variable_set[146] == 1) { return result;}
  result = C9;
  variable_set[146] = 1;
  return result;
}

ExcelValue eu_g4() {
  static ExcelValue result;
  if(variable_set[147] == 1) { return result;}
  result = C10;
  variable_set[147] = 1;
  return result;
}

ExcelValue eu_h4() {
  static ExcelValue result;
  if(variable_set[148] == 1) { return result;}
  result = C11;
  variable_set[148] = 1;
  return result;
}

ExcelValue eu_i4() {
  static ExcelValue result;
  if(variable_set[149] == 1) { return result;}
  result = C12;
  variable_set[149] = 1;
  return result;
}

ExcelValue eu_j4() {
  static ExcelValue result;
  if(variable_set[150] == 1) { return result;}
  result = C13;
  variable_set[150] = 1;
  return result;
}

ExcelValue eu_k4() {
  static ExcelValue result;
  if(variable_set[151] == 1) { return result;}
  result = C14;
  variable_set[151] = 1;
  return result;
}

ExcelValue eu_l4() {
  static ExcelValue result;
  if(variable_set[152] == 1) { return result;}
  result = C15;
  variable_set[152] = 1;
  return result;
}

ExcelValue eu_m4() {
  static ExcelValue result;
  if(variable_set[153] == 1) { return result;}
  result = C16;
  variable_set[153] = 1;
  return result;
}

ExcelValue eu_o4() {
  static ExcelValue result;
  if(variable_set[154] == 1) { return result;}
  result = C17;
  variable_set[154] = 1;
  return result;
}

ExcelValue eu_b5() {
  static ExcelValue result;
  if(variable_set[155] == 1) { return result;}
  result = C18;
  variable_set[155] = 1;
  return result;
}

ExcelValue eu_c5() {
  static ExcelValue result;
  if(variable_set[156] == 1) { return result;}
  result = C19;
  variable_set[156] = 1;
  return result;
}

ExcelValue eu_d5() {
  static ExcelValue result;
  if(variable_set[157] == 1) { return result;}
  result = C20;
  variable_set[157] = 1;
  return result;
}

ExcelValue eu_e5() {
  static ExcelValue result;
  if(variable_set[158] == 1) { return result;}
  result = C20;
  variable_set[158] = 1;
  return result;
}

ExcelValue eu_f5() {
  static ExcelValue result;
  if(variable_set[159] == 1) { return result;}
  result = C21;
  variable_set[159] = 1;
  return result;
}

ExcelValue eu_g5() {
  static ExcelValue result;
  if(variable_set[160] == 1) { return result;}
  result = C22;
  variable_set[160] = 1;
  return result;
}

ExcelValue eu_h5() {
  static ExcelValue result;
  if(variable_set[161] == 1) { return result;}
  result = C23;
  variable_set[161] = 1;
  return result;
}

ExcelValue eu_i5() {
  static ExcelValue result;
  if(variable_set[162] == 1) { return result;}
  result = multiply(C23,C24);
  variable_set[162] = 1;
  return result;
}

ExcelValue eu_j5() {
  static ExcelValue result;
  if(variable_set[163] == 1) { return result;}
  result = multiply(eu_i5(),C24);
  variable_set[163] = 1;
  return result;
}

ExcelValue eu_k5() {
  static ExcelValue result;
  if(variable_set[164] == 1) { return result;}
  result = multiply(eu_j5(),C24);
  variable_set[164] = 1;
  return result;
}

ExcelValue eu_l5() {
  static ExcelValue result;
  if(variable_set[165] == 1) { return result;}
  result = multiply(eu_k5(),C24);
  variable_set[165] = 1;
  return result;
}

ExcelValue eu_m5() {
  static ExcelValue result;
  if(variable_set[166] == 1) { return result;}
  result = multiply(eu_l5(),C24);
  variable_set[166] = 1;
  return result;
}

ExcelValue eu_n5() {
  static ExcelValue result;
  if(variable_set[167] == 1) { return result;}
  result = C25;
  variable_set[167] = 1;
  return result;
}

ExcelValue eu_o5() {
  static ExcelValue result;
  if(variable_set[168] == 1) { return result;}
  result = C26;
  variable_set[168] = 1;
  return result;
}

ExcelValue eu_b6() {
  static ExcelValue result;
  if(variable_set[169] == 1) { return result;}
  result = C27;
  variable_set[169] = 1;
  return result;
}

ExcelValue eu_c6() {
  static ExcelValue result;
  if(variable_set[170] == 1) { return result;}
  result = C28;
  variable_set[170] = 1;
  return result;
}

ExcelValue eu_f6() {
  static ExcelValue result;
  if(variable_set[171] == 1) { return result;}
  result = C29;
  variable_set[171] = 1;
  return result;
}

ExcelValue eu_g6() {
  static ExcelValue result;
  if(variable_set[172] == 1) { return result;}
  result = C30;
  variable_set[172] = 1;
  return result;
}

ExcelValue eu_h6() {
  static ExcelValue result;
  if(variable_set[173] == 1) { return result;}
  result = C31;
  variable_set[173] = 1;
  return result;
}

ExcelValue eu_i6() {
  static ExcelValue result;
  if(variable_set[174] == 1) { return result;}
  result = multiply(C31,C24);
  variable_set[174] = 1;
  return result;
}

ExcelValue eu_j6() {
  static ExcelValue result;
  if(variable_set[175] == 1) { return result;}
  result = multiply(eu_i6(),C24);
  variable_set[175] = 1;
  return result;
}

ExcelValue eu_k6() {
  static ExcelValue result;
  if(variable_set[176] == 1) { return result;}
  result = multiply(eu_j6(),C24);
  variable_set[176] = 1;
  return result;
}

ExcelValue eu_l6() {
  static ExcelValue result;
  if(variable_set[177] == 1) { return result;}
  result = multiply(eu_k6(),C24);
  variable_set[177] = 1;
  return result;
}

ExcelValue eu_m6() {
  static ExcelValue result;
  if(variable_set[178] == 1) { return result;}
  result = multiply(eu_l6(),C24);
  variable_set[178] = 1;
  return result;
}

ExcelValue eu_n6() {
  static ExcelValue result;
  if(variable_set[179] == 1) { return result;}
  result = C25;
  variable_set[179] = 1;
  return result;
}

ExcelValue eu_o6() {
  static ExcelValue result;
  if(variable_set[180] == 1) { return result;}
  result = C32;
  variable_set[180] = 1;
  return result;
}

ExcelValue eu_b7() {
  static ExcelValue result;
  if(variable_set[181] == 1) { return result;}
  result = C33;
  variable_set[181] = 1;
  return result;
}

ExcelValue eu_c7() {
  static ExcelValue result;
  if(variable_set[182] == 1) { return result;}
  result = C34;
  variable_set[182] = 1;
  return result;
}

ExcelValue eu_d7() {
  static ExcelValue result;
  if(variable_set[183] == 1) { return result;}
  result = C35;
  variable_set[183] = 1;
  return result;
}

ExcelValue eu_e7() {
  static ExcelValue result;
  if(variable_set[184] == 1) { return result;}
  result = C35;
  variable_set[184] = 1;
  return result;
}

ExcelValue eu_f7() {
  static ExcelValue result;
  if(variable_set[185] == 1) { return result;}
  result = C36;
  variable_set[185] = 1;
  return result;
}

ExcelValue eu_g7() {
  static ExcelValue result;
  if(variable_set[186] == 1) { return result;}
  result = C37;
  variable_set[186] = 1;
  return result;
}

ExcelValue eu_h7() {
  static ExcelValue result;
  if(variable_set[187] == 1) { return result;}
  result = C38;
  variable_set[187] = 1;
  return result;
}

ExcelValue eu_i7() {
  static ExcelValue result;
  if(variable_set[188] == 1) { return result;}
  result = multiply(C38,C24);
  variable_set[188] = 1;
  return result;
}

ExcelValue eu_j7() {
  static ExcelValue result;
  if(variable_set[189] == 1) { return result;}
  result = multiply(eu_i7(),C24);
  variable_set[189] = 1;
  return result;
}

ExcelValue eu_k7() {
  static ExcelValue result;
  if(variable_set[190] == 1) { return result;}
  result = multiply(eu_j7(),C24);
  variable_set[190] = 1;
  return result;
}

ExcelValue eu_l7() {
  static ExcelValue result;
  if(variable_set[191] == 1) { return result;}
  result = multiply(eu_k7(),C24);
  variable_set[191] = 1;
  return result;
}

ExcelValue eu_m7() {
  static ExcelValue result;
  if(variable_set[192] == 1) { return result;}
  result = multiply(eu_l7(),C24);
  variable_set[192] = 1;
  return result;
}

ExcelValue eu_n7() {
  static ExcelValue result;
  if(variable_set[193] == 1) { return result;}
  result = C25;
  variable_set[193] = 1;
  return result;
}

ExcelValue eu_o7() {
  static ExcelValue result;
  if(variable_set[194] == 1) { return result;}
  result = C32;
  variable_set[194] = 1;
  return result;
}

ExcelValue eu_b8() {
  static ExcelValue result;
  if(variable_set[195] == 1) { return result;}
  result = C39;
  variable_set[195] = 1;
  return result;
}

ExcelValue eu_f8() {
  static ExcelValue result;
  if(variable_set[196] == 1) { return result;}
  result = C40;
  variable_set[196] = 1;
  return result;
}

ExcelValue eu_g8() {
  static ExcelValue result;
  if(variable_set[197] == 1) { return result;}
  result = C41;
  variable_set[197] = 1;
  return result;
}

ExcelValue eu_h8() {
  static ExcelValue result;
  if(variable_set[198] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = C23;
  array1[1] = C31;
  array1[2] = C38;
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[198] = 1;
  return result;
}

ExcelValue eu_i8() {
  static ExcelValue result;
  if(variable_set[199] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_i5();
  array1[1] = eu_i6();
  array1[2] = eu_i7();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[199] = 1;
  return result;
}

ExcelValue eu_j8() {
  static ExcelValue result;
  if(variable_set[200] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_j5();
  array1[1] = eu_j6();
  array1[2] = eu_j7();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[200] = 1;
  return result;
}

ExcelValue eu_k8() {
  static ExcelValue result;
  if(variable_set[201] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_k5();
  array1[1] = eu_k6();
  array1[2] = eu_k7();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[201] = 1;
  return result;
}

ExcelValue eu_l8() {
  static ExcelValue result;
  if(variable_set[202] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_l5();
  array1[1] = eu_l6();
  array1[2] = eu_l7();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[202] = 1;
  return result;
}

ExcelValue eu_m8() {
  static ExcelValue result;
  if(variable_set[203] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = eu_m5();
  array1[1] = eu_m6();
  array1[2] = eu_m7();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[203] = 1;
  return result;
}

ExcelValue eu_n8() {
  static ExcelValue result;
  if(variable_set[204] == 1) { return result;}
  result = C25;
  variable_set[204] = 1;
  return result;
}

ExcelValue eu_b10() {
  static ExcelValue result;
  if(variable_set[205] == 1) { return result;}
  result = C42;
  variable_set[205] = 1;
  return result;
}

ExcelValue eu_f10() {
  static ExcelValue result;
  if(variable_set[206] == 1) { return result;}
  result = C9;
  variable_set[206] = 1;
  return result;
}

ExcelValue eu_g10() {
  static ExcelValue result;
  if(variable_set[207] == 1) { return result;}
  result = C10;
  variable_set[207] = 1;
  return result;
}

ExcelValue eu_h10() {
  static ExcelValue result;
  if(variable_set[208] == 1) { return result;}
  result = C11;
  variable_set[208] = 1;
  return result;
}

ExcelValue eu_i10() {
  static ExcelValue result;
  if(variable_set[209] == 1) { return result;}
  result = C12;
  variable_set[209] = 1;
  return result;
}

ExcelValue eu_j10() {
  static ExcelValue result;
  if(variable_set[210] == 1) { return result;}
  result = C13;
  variable_set[210] = 1;
  return result;
}

ExcelValue eu_k10() {
  static ExcelValue result;
  if(variable_set[211] == 1) { return result;}
  result = C14;
  variable_set[211] = 1;
  return result;
}

ExcelValue eu_l10() {
  static ExcelValue result;
  if(variable_set[212] == 1) { return result;}
  result = C15;
  variable_set[212] = 1;
  return result;
}

ExcelValue eu_m10() {
  static ExcelValue result;
  if(variable_set[213] == 1) { return result;}
  result = C16;
  variable_set[213] = 1;
  return result;
}

ExcelValue eu_b11() {
  static ExcelValue result;
  if(variable_set[214] == 1) { return result;}
  result = C18;
  variable_set[214] = 1;
  return result;
}

ExcelValue eu_f11() {
  static ExcelValue result;
  if(variable_set[215] == 1) { return result;}
  result = C43;
  variable_set[215] = 1;
  return result;
}

ExcelValue eu_g11() {
  static ExcelValue result;
  if(variable_set[216] == 1) { return result;}
  result = C43;
  variable_set[216] = 1;
  return result;
}

ExcelValue eu_h11() {
  static ExcelValue result;
  if(variable_set[217] == 1) { return result;}
  result = C43;
  variable_set[217] = 1;
  return result;
}

ExcelValue eu_i11() {
  static ExcelValue result;
  if(variable_set[218] == 1) { return result;}
  result = C43;
  variable_set[218] = 1;
  return result;
}

ExcelValue eu_j11() {
  static ExcelValue result;
  if(variable_set[219] == 1) { return result;}
  result = C43;
  variable_set[219] = 1;
  return result;
}

ExcelValue eu_k11() {
  static ExcelValue result;
  if(variable_set[220] == 1) { return result;}
  result = C43;
  variable_set[220] = 1;
  return result;
}

ExcelValue eu_l11() {
  static ExcelValue result;
  if(variable_set[221] == 1) { return result;}
  result = C43;
  variable_set[221] = 1;
  return result;
}

ExcelValue eu_m11() {
  static ExcelValue result;
  if(variable_set[222] == 1) { return result;}
  result = C43;
  variable_set[222] = 1;
  return result;
}

ExcelValue eu_n11() {
  static ExcelValue result;
  if(variable_set[223] == 1) { return result;}
  result = C44;
  variable_set[223] = 1;
  return result;
}

ExcelValue eu_b12() {
  static ExcelValue result;
  if(variable_set[224] == 1) { return result;}
  result = C27;
  variable_set[224] = 1;
  return result;
}

ExcelValue eu_f12() {
  static ExcelValue result;
  if(variable_set[225] == 1) { return result;}
  result = C45;
  variable_set[225] = 1;
  return result;
}

ExcelValue eu_g12() {
  static ExcelValue result;
  if(variable_set[226] == 1) { return result;}
  result = C46;
  variable_set[226] = 1;
  return result;
}

ExcelValue eu_h12() {
  static ExcelValue result;
  if(variable_set[227] == 1) { return result;}
  result = C47;
  variable_set[227] = 1;
  return result;
}

ExcelValue eu_i12() {
  static ExcelValue result;
  if(variable_set[228] == 1) { return result;}
  result = C48;
  variable_set[228] = 1;
  return result;
}

ExcelValue eu_j12() {
  static ExcelValue result;
  if(variable_set[229] == 1) { return result;}
  result = add(C48,C49);
  variable_set[229] = 1;
  return result;
}

ExcelValue eu_k12() {
  static ExcelValue result;
  if(variable_set[230] == 1) { return result;}
  result = add(eu_j12(),C49);
  variable_set[230] = 1;
  return result;
}

ExcelValue eu_l12() {
  static ExcelValue result;
  if(variable_set[231] == 1) { return result;}
  result = add(eu_k12(),C49);
  variable_set[231] = 1;
  return result;
}

ExcelValue eu_m12() {
  static ExcelValue result;
  if(variable_set[232] == 1) { return result;}
  result = add(eu_l12(),C49);
  variable_set[232] = 1;
  return result;
}

ExcelValue eu_o12() {
  static ExcelValue result;
  if(variable_set[233] == 1) { return result;}
  result = C49;
  variable_set[233] = 1;
  return result;
}

ExcelValue eu_b13() {
  static ExcelValue result;
  if(variable_set[234] == 1) { return result;}
  result = C50;
  variable_set[234] = 1;
  return result;
}

ExcelValue eu_f13() {
  static ExcelValue result;
  if(variable_set[235] == 1) { return result;}
  result = C51;
  variable_set[235] = 1;
  return result;
}

ExcelValue eu_g13() {
  static ExcelValue result;
  if(variable_set[236] == 1) { return result;}
  result = C52;
  variable_set[236] = 1;
  return result;
}

ExcelValue eu_h13() {
  static ExcelValue result;
  if(variable_set[237] == 1) { return result;}
  result = C53;
  variable_set[237] = 1;
  return result;
}

ExcelValue eu_i13() {
  static ExcelValue result;
  if(variable_set[238] == 1) { return result;}
  result = add(C53,C54);
  variable_set[238] = 1;
  return result;
}

ExcelValue eu_j13() {
  static ExcelValue result;
  if(variable_set[239] == 1) { return result;}
  result = add(eu_i13(),C54);
  variable_set[239] = 1;
  return result;
}

ExcelValue eu_k13() {
  static ExcelValue result;
  if(variable_set[240] == 1) { return result;}
  result = add(eu_j13(),C54);
  variable_set[240] = 1;
  return result;
}

ExcelValue eu_l13() {
  static ExcelValue result;
  if(variable_set[241] == 1) { return result;}
  result = add(eu_k13(),C54);
  variable_set[241] = 1;
  return result;
}

ExcelValue eu_m13() {
  static ExcelValue result;
  if(variable_set[242] == 1) { return result;}
  result = C43;
  variable_set[242] = 1;
  return result;
}

ExcelValue eu_n13() {
  static ExcelValue result;
  if(variable_set[243] == 1) { return result;}
  result = C44;
  variable_set[243] = 1;
  return result;
}

ExcelValue eu_o13() {
  static ExcelValue result;
  if(variable_set[244] == 1) { return result;}
  result = C54;
  variable_set[244] = 1;
  return result;
}

ExcelValue eu_b14() {
  static ExcelValue result;
  if(variable_set[245] == 1) { return result;}
  result = C55;
  variable_set[245] = 1;
  return result;
}

ExcelValue eu_f14() {
  static ExcelValue result;
  if(variable_set[246] == 1) { return result;}
  result = C56;
  variable_set[246] = 1;
  return result;
}

ExcelValue eu_g14() {
  static ExcelValue result;
  if(variable_set[247] == 1) { return result;}
  result = C57;
  variable_set[247] = 1;
  return result;
}

ExcelValue eu_h14() {
  static ExcelValue result;
  if(variable_set[248] == 1) { return result;}
  result = _common0();
  variable_set[248] = 1;
  return result;
}

ExcelValue eu_i14() {
  static ExcelValue result;
  if(variable_set[249] == 1) { return result;}
  result = _common6();
  variable_set[249] = 1;
  return result;
}

ExcelValue eu_j14() {
  static ExcelValue result;
  if(variable_set[250] == 1) { return result;}
  result = _common12();
  variable_set[250] = 1;
  return result;
}

ExcelValue eu_k14() {
  static ExcelValue result;
  if(variable_set[251] == 1) { return result;}
  result = _common18();
  variable_set[251] = 1;
  return result;
}

ExcelValue eu_l14() {
  static ExcelValue result;
  if(variable_set[252] == 1) { return result;}
  result = _common24();
  variable_set[252] = 1;
  return result;
}

ExcelValue eu_m14() {
  static ExcelValue result;
  if(variable_set[253] == 1) { return result;}
  result = _common30();
  variable_set[253] = 1;
  return result;
}

ExcelValue eu_b15() {
  static ExcelValue result;
  if(variable_set[254] == 1) { return result;}
  result = C58;
  variable_set[254] = 1;
  return result;
}

ExcelValue eu_f15() {
  static ExcelValue result;
  if(variable_set[255] == 1) { return result;}
  result = C59;
  variable_set[255] = 1;
  return result;
}

ExcelValue eu_g15() {
  static ExcelValue result;
  if(variable_set[256] == 1) { return result;}
  result = divide(C60,C41);
  variable_set[256] = 1;
  return result;
}

ExcelValue eu_h15() {
  static ExcelValue result;
  if(variable_set[257] == 1) { return result;}
  result = divide(add(add(multiply(C43,C23),multiply(C47,C31)),multiply(C53,C38)),eu_h8());
  variable_set[257] = 1;
  return result;
}

ExcelValue eu_i15() {
  static ExcelValue result;
  if(variable_set[258] == 1) { return result;}
  result = divide(add(add(multiply(C43,eu_i5()),multiply(C48,eu_i6())),multiply(eu_i13(),eu_i7())),eu_i8());
  variable_set[258] = 1;
  return result;
}

ExcelValue eu_j15() {
  static ExcelValue result;
  if(variable_set[259] == 1) { return result;}
  result = divide(add(add(multiply(C43,eu_j5()),multiply(eu_j12(),eu_j6())),multiply(eu_j13(),eu_j7())),eu_j8());
  variable_set[259] = 1;
  return result;
}

ExcelValue eu_k15() {
  static ExcelValue result;
  if(variable_set[260] == 1) { return result;}
  result = divide(add(add(multiply(C43,eu_k5()),multiply(eu_k12(),eu_k6())),multiply(eu_k13(),eu_k7())),eu_k8());
  variable_set[260] = 1;
  return result;
}

ExcelValue eu_l15() {
  static ExcelValue result;
  if(variable_set[261] == 1) { return result;}
  result = divide(add(add(multiply(C43,eu_l5()),multiply(eu_l12(),eu_l6())),multiply(eu_l13(),eu_l7())),eu_l8());
  variable_set[261] = 1;
  return result;
}

ExcelValue eu_m15() {
  static ExcelValue result;
  if(variable_set[262] == 1) { return result;}
  result = divide(add(add(multiply(C43,eu_m5()),multiply(eu_m12(),eu_m6())),multiply(C43,eu_m7())),eu_m8());
  variable_set[262] = 1;
  return result;
}

ExcelValue eu_n15() {
  static ExcelValue result;
  if(variable_set[263] == 1) { return result;}
  result = C44;
  variable_set[263] = 1;
  return result;
}

ExcelValue eu_b17() {
  static ExcelValue result;
  if(variable_set[264] == 1) { return result;}
  result = C61;
  variable_set[264] = 1;
  return result;
}

ExcelValue eu_f17() {
  static ExcelValue result;
  if(variable_set[265] == 1) { return result;}
  result = C62;
  variable_set[265] = 1;
  return result;
}

ExcelValue eu_g17() {
  static ExcelValue result;
  if(variable_set[266] == 1) { return result;}
  result = subtract(C63,C57);
  variable_set[266] = 1;
  return result;
}

ExcelValue eu_h17() {
  static ExcelValue result;
  if(variable_set[267] == 1) { return result;}
  result = subtract(C63,_common0());
  variable_set[267] = 1;
  return result;
}

ExcelValue eu_i17() {
  static ExcelValue result;
  if(variable_set[268] == 1) { return result;}
  result = subtract(C63,_common6());
  variable_set[268] = 1;
  return result;
}

ExcelValue eu_j17() {
  static ExcelValue result;
  if(variable_set[269] == 1) { return result;}
  result = subtract(C63,_common12());
  variable_set[269] = 1;
  return result;
}

ExcelValue eu_k17() {
  static ExcelValue result;
  if(variable_set[270] == 1) { return result;}
  result = subtract(C63,_common18());
  variable_set[270] = 1;
  return result;
}

ExcelValue eu_l17() {
  static ExcelValue result;
  if(variable_set[271] == 1) { return result;}
  result = subtract(C63,_common24());
  variable_set[271] = 1;
  return result;
}

ExcelValue eu_m17() {
  static ExcelValue result;
  if(variable_set[272] == 1) { return result;}
  result = subtract(C63,_common30());
  variable_set[272] = 1;
  return result;
}

ExcelValue eu_b18() {
  static ExcelValue result;
  if(variable_set[273] == 1) { return result;}
  result = C64;
  variable_set[273] = 1;
  return result;
}

ExcelValue eu_d18() {
  static ExcelValue result;
  if(variable_set[274] == 1) { return result;}
  result = C65;
  variable_set[274] = 1;
  return result;
}

ExcelValue eu_f18() {
  static ExcelValue result;
  if(variable_set[275] == 1) { return result;}
  result = C9;
  variable_set[275] = 1;
  return result;
}

ExcelValue eu_g18() {
  static ExcelValue result;
  if(variable_set[276] == 1) { return result;}
  result = C10;
  variable_set[276] = 1;
  return result;
}

ExcelValue eu_h18() {
  static ExcelValue result;
  if(variable_set[277] == 1) { return result;}
  result = C11;
  variable_set[277] = 1;
  return result;
}

ExcelValue eu_i18() {
  static ExcelValue result;
  if(variable_set[278] == 1) { return result;}
  result = C12;
  variable_set[278] = 1;
  return result;
}

ExcelValue eu_j18() {
  static ExcelValue result;
  if(variable_set[279] == 1) { return result;}
  result = C13;
  variable_set[279] = 1;
  return result;
}

ExcelValue eu_k18() {
  static ExcelValue result;
  if(variable_set[280] == 1) { return result;}
  result = C14;
  variable_set[280] = 1;
  return result;
}

ExcelValue eu_l18() {
  static ExcelValue result;
  if(variable_set[281] == 1) { return result;}
  result = C15;
  variable_set[281] = 1;
  return result;
}

ExcelValue eu_m18() {
  static ExcelValue result;
  if(variable_set[282] == 1) { return result;}
  result = C16;
  variable_set[282] = 1;
  return result;
}

ExcelValue eu_b19() {
  static ExcelValue result;
  if(variable_set[283] == 1) { return result;}
  result = C64;
  variable_set[283] = 1;
  return result;
}

ExcelValue eu_d19() {
  static ExcelValue result;
  if(variable_set[284] == 1) { return result;}
  result = C66;
  variable_set[284] = 1;
  return result;
}

ExcelValue eu_f19() {
  static ExcelValue result;
  if(variable_set[285] == 1) { return result;}
  result = C67;
  variable_set[285] = 1;
  return result;
}

ExcelValue eu_g19() {
  static ExcelValue result;
  if(variable_set[286] == 1) { return result;}
  result = C68;
  variable_set[286] = 1;
  return result;
}

ExcelValue eu_h19() {
  static ExcelValue result;
  if(variable_set[287] == 1) { return result;}
  result = C69;
  variable_set[287] = 1;
  return result;
}

ExcelValue eu_i19() {
  static ExcelValue result;
  if(variable_set[288] == 1) { return result;}
  result = multiply(_common36(),C70);
  variable_set[288] = 1;
  return result;
}

ExcelValue eu_j19() {
  static ExcelValue result;
  if(variable_set[289] == 1) { return result;}
  result = multiply(_common37(),C70);
  variable_set[289] = 1;
  return result;
}

ExcelValue eu_k19() {
  static ExcelValue result;
  if(variable_set[290] == 1) { return result;}
  result = multiply(_common38(),C70);
  variable_set[290] = 1;
  return result;
}

ExcelValue eu_l19() {
  static ExcelValue result;
  if(variable_set[291] == 1) { return result;}
  result = multiply(_common39(),C70);
  variable_set[291] = 1;
  return result;
}

ExcelValue eu_m19() {
  static ExcelValue result;
  if(variable_set[292] == 1) { return result;}
  result = multiply(_common40(),C70);
  variable_set[292] = 1;
  return result;
}

ExcelValue eu_o19() {
  static ExcelValue result;
  if(variable_set[293] == 1) { return result;}
  result = C71;
  variable_set[293] = 1;
  return result;
}

ExcelValue eu_p19() {
  static ExcelValue result;
  if(variable_set[294] == 1) { return result;}
  result = C72;
  variable_set[294] = 1;
  return result;
}

ExcelValue eu_b20() {
  static ExcelValue result;
  if(variable_set[295] == 1) { return result;}
  result = C73;
  variable_set[295] = 1;
  return result;
}

ExcelValue eu_f20() {
  static ExcelValue result;
  if(variable_set[296] == 1) { return result;}
  result = C74;
  variable_set[296] = 1;
  return result;
}

ExcelValue eu_g20() {
  static ExcelValue result;
  if(variable_set[297] == 1) { return result;}
  result = _common41();
  variable_set[297] = 1;
  return result;
}

ExcelValue eu_h20() {
  static ExcelValue result;
  if(variable_set[298] == 1) { return result;}
  result = _common43();
  variable_set[298] = 1;
  return result;
}

ExcelValue eu_i20() {
  static ExcelValue result;
  if(variable_set[299] == 1) { return result;}
  result = _common47();
  variable_set[299] = 1;
  return result;
}

ExcelValue eu_j20() {
  static ExcelValue result;
  if(variable_set[300] == 1) { return result;}
  result = _common52();
  variable_set[300] = 1;
  return result;
}

ExcelValue eu_k20() {
  static ExcelValue result;
  if(variable_set[301] == 1) { return result;}
  result = _common57();
  variable_set[301] = 1;
  return result;
}

ExcelValue eu_l20() {
  static ExcelValue result;
  if(variable_set[302] == 1) { return result;}
  result = _common62();
  variable_set[302] = 1;
  return result;
}

ExcelValue eu_m20() {
  static ExcelValue result;
  if(variable_set[303] == 1) { return result;}
  result = _common67();
  variable_set[303] = 1;
  return result;
}

ExcelValue eu_b21() {
  static ExcelValue result;
  if(variable_set[304] == 1) { return result;}
  result = C75;
  variable_set[304] = 1;
  return result;
}

ExcelValue eu_f21() {
  static ExcelValue result;
  if(variable_set[305] == 1) { return result;}
  result = _common71();
  variable_set[305] = 1;
  return result;
}

ExcelValue eu_g21() {
  static ExcelValue result;
  if(variable_set[306] == 1) { return result;}
  result = _common72();
  variable_set[306] = 1;
  return result;
}

ExcelValue eu_h21() {
  static ExcelValue result;
  if(variable_set[307] == 1) { return result;}
  result = _common73();
  variable_set[307] = 1;
  return result;
}

ExcelValue eu_i21() {
  static ExcelValue result;
  if(variable_set[308] == 1) { return result;}
  result = _common74();
  variable_set[308] = 1;
  return result;
}

ExcelValue eu_j21() {
  static ExcelValue result;
  if(variable_set[309] == 1) { return result;}
  result = _common75();
  variable_set[309] = 1;
  return result;
}

ExcelValue eu_k21() {
  static ExcelValue result;
  if(variable_set[310] == 1) { return result;}
  result = _common76();
  variable_set[310] = 1;
  return result;
}

ExcelValue eu_l21() {
  static ExcelValue result;
  if(variable_set[311] == 1) { return result;}
  result = _common77();
  variable_set[311] = 1;
  return result;
}

ExcelValue eu_m21() {
  static ExcelValue result;
  if(variable_set[312] == 1) { return result;}
  result = _common78();
  variable_set[312] = 1;
  return result;
}

ExcelValue eu_b23() {
  static ExcelValue result;
  if(variable_set[313] == 1) { return result;}
  result = C76;
  variable_set[313] = 1;
  return result;
}

ExcelValue eu_f23() {
  static ExcelValue result;
  if(variable_set[314] == 1) { return result;}
  result = C9;
  variable_set[314] = 1;
  return result;
}

ExcelValue eu_g23() {
  static ExcelValue result;
  if(variable_set[315] == 1) { return result;}
  result = C10;
  variable_set[315] = 1;
  return result;
}

ExcelValue eu_h23() {
  static ExcelValue result;
  if(variable_set[316] == 1) { return result;}
  result = C11;
  variable_set[316] = 1;
  return result;
}

ExcelValue eu_i23() {
  static ExcelValue result;
  if(variable_set[317] == 1) { return result;}
  result = C12;
  variable_set[317] = 1;
  return result;
}

ExcelValue eu_j23() {
  static ExcelValue result;
  if(variable_set[318] == 1) { return result;}
  result = C13;
  variable_set[318] = 1;
  return result;
}

ExcelValue eu_k23() {
  static ExcelValue result;
  if(variable_set[319] == 1) { return result;}
  result = C14;
  variable_set[319] = 1;
  return result;
}

ExcelValue eu_l23() {
  static ExcelValue result;
  if(variable_set[320] == 1) { return result;}
  result = C15;
  variable_set[320] = 1;
  return result;
}

ExcelValue eu_m23() {
  static ExcelValue result;
  if(variable_set[321] == 1) { return result;}
  result = C16;
  variable_set[321] = 1;
  return result;
}

ExcelValue eu_b24() {
  static ExcelValue result;
  if(variable_set[322] == 1) { return result;}
  result = C77;
  variable_set[322] = 1;
  return result;
}

ExcelValue eu_f24() {
  static ExcelValue result;
  if(variable_set[323] == 1) { return result;}
  result = C78;
  variable_set[323] = 1;
  return result;
}

ExcelValue eu_g24() {
  static ExcelValue result;
  if(variable_set[324] == 1) { return result;}
  result = C79;
  variable_set[324] = 1;
  return result;
}

ExcelValue eu_h24() {
  static ExcelValue result;
  if(variable_set[325] == 1) { return result;}
  result = C80;
  variable_set[325] = 1;
  return result;
}

ExcelValue eu_i24() {
  static ExcelValue result;
  if(variable_set[326] == 1) { return result;}
  result = C81;
  variable_set[326] = 1;
  return result;
}

ExcelValue eu_j24() {
  static ExcelValue result;
  if(variable_set[327] == 1) { return result;}
  result = multiply(_common80(),C82);
  variable_set[327] = 1;
  return result;
}

ExcelValue eu_k24() {
  static ExcelValue result;
  if(variable_set[328] == 1) { return result;}
  result = multiply(_common81(),C82);
  variable_set[328] = 1;
  return result;
}

ExcelValue eu_l24() {
  static ExcelValue result;
  if(variable_set[329] == 1) { return result;}
  result = multiply(_common82(),C82);
  variable_set[329] = 1;
  return result;
}

ExcelValue eu_m24() {
  static ExcelValue result;
  if(variable_set[330] == 1) { return result;}
  result = multiply(_common83(),C82);
  variable_set[330] = 1;
  return result;
}

ExcelValue eu_n24() {
  static ExcelValue result;
  if(variable_set[331] == 1) { return result;}
  result = C83;
  variable_set[331] = 1;
  return result;
}

ExcelValue eu_o24() {
  static ExcelValue result;
  if(variable_set[332] == 1) { return result;}
  result = C84;
  variable_set[332] = 1;
  return result;
}

ExcelValue eu_b25() {
  static ExcelValue result;
  if(variable_set[333] == 1) { return result;}
  result = C85;
  variable_set[333] = 1;
  return result;
}

ExcelValue eu_f25() {
  static ExcelValue result;
  if(variable_set[334] == 1) { return result;}
  result = multiply(C78,_common71());
  variable_set[334] = 1;
  return result;
}

ExcelValue eu_g25() {
  static ExcelValue result;
  if(variable_set[335] == 1) { return result;}
  result = multiply(C79,_common72());
  variable_set[335] = 1;
  return result;
}

ExcelValue eu_h25() {
  static ExcelValue result;
  if(variable_set[336] == 1) { return result;}
  result = multiply(C80,_common73());
  variable_set[336] = 1;
  return result;
}

ExcelValue eu_i25() {
  static ExcelValue result;
  if(variable_set[337] == 1) { return result;}
  result = multiply(_common80(),_common74());
  variable_set[337] = 1;
  return result;
}

ExcelValue eu_j25() {
  static ExcelValue result;
  if(variable_set[338] == 1) { return result;}
  result = multiply(_common81(),_common75());
  variable_set[338] = 1;
  return result;
}

ExcelValue eu_k25() {
  static ExcelValue result;
  if(variable_set[339] == 1) { return result;}
  result = multiply(_common82(),_common76());
  variable_set[339] = 1;
  return result;
}

ExcelValue eu_l25() {
  static ExcelValue result;
  if(variable_set[340] == 1) { return result;}
  result = multiply(_common83(),_common77());
  variable_set[340] = 1;
  return result;
}

ExcelValue eu_m25() {
  static ExcelValue result;
  if(variable_set[341] == 1) { return result;}
  result = multiply(eu_m24(),_common78());
  variable_set[341] = 1;
  return result;
}

ExcelValue eu_b27() {
  static ExcelValue result;
  if(variable_set[342] == 1) { return result;}
  result = C86;
  variable_set[342] = 1;
  return result;
}

ExcelValue eu_f27() {
  static ExcelValue result;
  if(variable_set[343] == 1) { return result;}
  result = C9;
  variable_set[343] = 1;
  return result;
}

ExcelValue eu_g27() {
  static ExcelValue result;
  if(variable_set[344] == 1) { return result;}
  result = C10;
  variable_set[344] = 1;
  return result;
}

ExcelValue eu_h27() {
  static ExcelValue result;
  if(variable_set[345] == 1) { return result;}
  result = C11;
  variable_set[345] = 1;
  return result;
}

ExcelValue eu_i27() {
  static ExcelValue result;
  if(variable_set[346] == 1) { return result;}
  result = C12;
  variable_set[346] = 1;
  return result;
}

ExcelValue eu_j27() {
  static ExcelValue result;
  if(variable_set[347] == 1) { return result;}
  result = C13;
  variable_set[347] = 1;
  return result;
}

ExcelValue eu_k27() {
  static ExcelValue result;
  if(variable_set[348] == 1) { return result;}
  result = C14;
  variable_set[348] = 1;
  return result;
}

ExcelValue eu_l27() {
  static ExcelValue result;
  if(variable_set[349] == 1) { return result;}
  result = C15;
  variable_set[349] = 1;
  return result;
}

ExcelValue eu_m27() {
  static ExcelValue result;
  if(variable_set[350] == 1) { return result;}
  result = C16;
  variable_set[350] = 1;
  return result;
}

ExcelValue eu_b28() {
  static ExcelValue result;
  if(variable_set[351] == 1) { return result;}
  result = C18;
  variable_set[351] = 1;
  return result;
}

ExcelValue eu_f28() {
  static ExcelValue result;
  if(variable_set[352] == 1) { return result;}
  result = C87;
  variable_set[352] = 1;
  return result;
}

ExcelValue eu_g28() {
  static ExcelValue result;
  if(variable_set[353] == 1) { return result;}
  result = multiply(C22,_common84());
  variable_set[353] = 1;
  return result;
}

ExcelValue eu_h28() {
  static ExcelValue result;
  if(variable_set[354] == 1) { return result;}
  result = multiply(multiply(C23,C63),_common85());
  variable_set[354] = 1;
  return result;
}

ExcelValue eu_i28() {
  static ExcelValue result;
  if(variable_set[355] == 1) { return result;}
  result = multiply(multiply(eu_i5(),C63),_common86());
  variable_set[355] = 1;
  return result;
}

ExcelValue eu_j28() {
  static ExcelValue result;
  if(variable_set[356] == 1) { return result;}
  result = multiply(multiply(eu_j5(),C63),_common87());
  variable_set[356] = 1;
  return result;
}

ExcelValue eu_k28() {
  static ExcelValue result;
  if(variable_set[357] == 1) { return result;}
  result = multiply(multiply(eu_k5(),C63),_common88());
  variable_set[357] = 1;
  return result;
}

ExcelValue eu_l28() {
  static ExcelValue result;
  if(variable_set[358] == 1) { return result;}
  result = multiply(multiply(eu_l5(),C63),_common89());
  variable_set[358] = 1;
  return result;
}

ExcelValue eu_m28() {
  static ExcelValue result;
  if(variable_set[359] == 1) { return result;}
  result = multiply(multiply(eu_m5(),C63),_common90());
  variable_set[359] = 1;
  return result;
}

ExcelValue eu_b29() {
  static ExcelValue result;
  if(variable_set[360] == 1) { return result;}
  result = C27;
  variable_set[360] = 1;
  return result;
}

ExcelValue eu_f29() {
  static ExcelValue result;
  if(variable_set[361] == 1) { return result;}
  result = C88;
  variable_set[361] = 1;
  return result;
}

ExcelValue eu_g29() {
  static ExcelValue result;
  if(variable_set[362] == 1) { return result;}
  result = multiply(C89,_common84());
  variable_set[362] = 1;
  return result;
}

ExcelValue eu_h29() {
  static ExcelValue result;
  if(variable_set[363] == 1) { return result;}
  result = multiply(multiply(C31,C90),_common85());
  variable_set[363] = 1;
  return result;
}

ExcelValue eu_i29() {
  static ExcelValue result;
  if(variable_set[364] == 1) { return result;}
  result = multiply(multiply(eu_i6(),subtract(C63,C48)),_common86());
  variable_set[364] = 1;
  return result;
}

ExcelValue eu_j29() {
  static ExcelValue result;
  if(variable_set[365] == 1) { return result;}
  result = multiply(multiply(eu_j6(),subtract(C63,eu_j12())),_common87());
  variable_set[365] = 1;
  return result;
}

ExcelValue eu_k29() {
  static ExcelValue result;
  if(variable_set[366] == 1) { return result;}
  result = multiply(multiply(eu_k6(),subtract(C63,eu_k12())),_common88());
  variable_set[366] = 1;
  return result;
}

ExcelValue eu_l29() {
  static ExcelValue result;
  if(variable_set[367] == 1) { return result;}
  result = multiply(multiply(eu_l6(),subtract(C63,eu_l12())),_common89());
  variable_set[367] = 1;
  return result;
}

ExcelValue eu_m29() {
  static ExcelValue result;
  if(variable_set[368] == 1) { return result;}
  result = multiply(multiply(eu_m6(),subtract(C63,eu_m12())),_common90());
  variable_set[368] = 1;
  return result;
}

ExcelValue eu_b30() {
  static ExcelValue result;
  if(variable_set[369] == 1) { return result;}
  result = C50;
  variable_set[369] = 1;
  return result;
}

ExcelValue eu_f30() {
  static ExcelValue result;
  if(variable_set[370] == 1) { return result;}
  result = C91;
  variable_set[370] = 1;
  return result;
}

ExcelValue eu_g30() {
  static ExcelValue result;
  if(variable_set[371] == 1) { return result;}
  result = multiply(C92,_common84());
  variable_set[371] = 1;
  return result;
}

ExcelValue eu_h30() {
  static ExcelValue result;
  if(variable_set[372] == 1) { return result;}
  result = multiply(multiply(C38,subtract(C63,C53)),_common85());
  variable_set[372] = 1;
  return result;
}

ExcelValue eu_i30() {
  static ExcelValue result;
  if(variable_set[373] == 1) { return result;}
  result = multiply(multiply(eu_i7(),subtract(C63,eu_i13())),_common86());
  variable_set[373] = 1;
  return result;
}

ExcelValue eu_j30() {
  static ExcelValue result;
  if(variable_set[374] == 1) { return result;}
  result = multiply(multiply(eu_j7(),subtract(C63,eu_j13())),_common87());
  variable_set[374] = 1;
  return result;
}

ExcelValue eu_k30() {
  static ExcelValue result;
  if(variable_set[375] == 1) { return result;}
  result = multiply(multiply(eu_k7(),subtract(C63,eu_k13())),_common88());
  variable_set[375] = 1;
  return result;
}

ExcelValue eu_l30() {
  static ExcelValue result;
  if(variable_set[376] == 1) { return result;}
  result = multiply(multiply(eu_l7(),subtract(C63,eu_l13())),_common89());
  variable_set[376] = 1;
  return result;
}

ExcelValue eu_m30() {
  static ExcelValue result;
  if(variable_set[377] == 1) { return result;}
  result = multiply(multiply(eu_m7(),C63),_common90());
  variable_set[377] = 1;
  return result;
}

ExcelValue eu_b31() {
  static ExcelValue result;
  if(variable_set[378] == 1) { return result;}
  result = C39;
  variable_set[378] = 1;
  return result;
}

ExcelValue eu_f31() {
  static ExcelValue result;
  if(variable_set[379] == 1) { return result;}
  result = _common91();
  variable_set[379] = 1;
  return result;
}

ExcelValue eu_g31() {
  static ExcelValue result;
  if(variable_set[380] == 1) { return result;}
  result = _common94();
  variable_set[380] = 1;
  return result;
}

ExcelValue eu_h31() {
  static ExcelValue result;
  if(variable_set[381] == 1) { return result;}
  result = _common97();
  variable_set[381] = 1;
  return result;
}

ExcelValue eu_i31() {
  static ExcelValue result;
  if(variable_set[382] == 1) { return result;}
  result = _common100();
  variable_set[382] = 1;
  return result;
}

ExcelValue eu_j31() {
  static ExcelValue result;
  if(variable_set[383] == 1) { return result;}
  result = _common103();
  variable_set[383] = 1;
  return result;
}

ExcelValue eu_k31() {
  static ExcelValue result;
  if(variable_set[384] == 1) { return result;}
  result = _common106();
  variable_set[384] = 1;
  return result;
}

ExcelValue eu_l31() {
  static ExcelValue result;
  if(variable_set[385] == 1) { return result;}
  result = _common109();
  variable_set[385] = 1;
  return result;
}

ExcelValue eu_m31() {
  static ExcelValue result;
  if(variable_set[386] == 1) { return result;}
  result = _common112();
  variable_set[386] = 1;
  return result;
}

ExcelValue eu_b33() {
  static ExcelValue result;
  if(variable_set[387] == 1) { return result;}
  result = C93;
  variable_set[387] = 1;
  return result;
}

ExcelValue eu_b34() {
  static ExcelValue result;
  if(variable_set[388] == 1) { return result;}
  result = C18;
  variable_set[388] = 1;
  return result;
}

ExcelValue eu_f34() {
  static ExcelValue result;
  if(variable_set[389] == 1) { return result;}
  result = multiply(C87,C78);
  variable_set[389] = 1;
  return result;
}

ExcelValue eu_g34() {
  static ExcelValue result;
  if(variable_set[390] == 1) { return result;}
  result = multiply(eu_g28(),C79);
  variable_set[390] = 1;
  return result;
}

ExcelValue eu_h34() {
  static ExcelValue result;
  if(variable_set[391] == 1) { return result;}
  result = multiply(eu_h28(),C80);
  variable_set[391] = 1;
  return result;
}

ExcelValue eu_i34() {
  static ExcelValue result;
  if(variable_set[392] == 1) { return result;}
  result = multiply(eu_i28(),C81);
  variable_set[392] = 1;
  return result;
}

ExcelValue eu_j34() {
  static ExcelValue result;
  if(variable_set[393] == 1) { return result;}
  result = multiply(eu_j28(),eu_j24());
  variable_set[393] = 1;
  return result;
}

ExcelValue eu_k34() {
  static ExcelValue result;
  if(variable_set[394] == 1) { return result;}
  result = multiply(eu_k28(),eu_k24());
  variable_set[394] = 1;
  return result;
}

ExcelValue eu_l34() {
  static ExcelValue result;
  if(variable_set[395] == 1) { return result;}
  result = multiply(eu_l28(),eu_l24());
  variable_set[395] = 1;
  return result;
}

ExcelValue eu_m34() {
  static ExcelValue result;
  if(variable_set[396] == 1) { return result;}
  result = multiply(eu_m28(),eu_m24());
  variable_set[396] = 1;
  return result;
}

ExcelValue eu_b35() {
  static ExcelValue result;
  if(variable_set[397] == 1) { return result;}
  result = C33;
  variable_set[397] = 1;
  return result;
}

ExcelValue eu_f35() {
  static ExcelValue result;
  if(variable_set[398] == 1) { return result;}
  result = multiply(add(C88,C91),C78);
  variable_set[398] = 1;
  return result;
}

ExcelValue eu_g35() {
  static ExcelValue result;
  if(variable_set[399] == 1) { return result;}
  result = multiply(add(eu_g29(),eu_g30()),C79);
  variable_set[399] = 1;
  return result;
}

ExcelValue eu_h35() {
  static ExcelValue result;
  if(variable_set[400] == 1) { return result;}
  result = multiply(add(eu_h29(),eu_h30()),C80);
  variable_set[400] = 1;
  return result;
}

ExcelValue eu_i35() {
  static ExcelValue result;
  if(variable_set[401] == 1) { return result;}
  result = multiply(add(eu_i29(),eu_i30()),C81);
  variable_set[401] = 1;
  return result;
}

ExcelValue eu_j35() {
  static ExcelValue result;
  if(variable_set[402] == 1) { return result;}
  result = multiply(add(eu_j29(),eu_j30()),eu_j24());
  variable_set[402] = 1;
  return result;
}

ExcelValue eu_k35() {
  static ExcelValue result;
  if(variable_set[403] == 1) { return result;}
  result = multiply(add(eu_k29(),eu_k30()),eu_k24());
  variable_set[403] = 1;
  return result;
}

ExcelValue eu_l35() {
  static ExcelValue result;
  if(variable_set[404] == 1) { return result;}
  result = multiply(add(eu_l29(),eu_l30()),eu_l24());
  variable_set[404] = 1;
  return result;
}

ExcelValue eu_m35() {
  static ExcelValue result;
  if(variable_set[405] == 1) { return result;}
  result = multiply(add(eu_m29(),eu_m30()),eu_m24());
  variable_set[405] = 1;
  return result;
}

ExcelValue eu_b39() {
  static ExcelValue result;
  if(variable_set[406] == 1) { return result;}
  result = C94;
  variable_set[406] = 1;
  return result;
}

ExcelValue eu_f39() {
  static ExcelValue result;
  if(variable_set[407] == 1) { return result;}
  result = C95;
  variable_set[407] = 1;
  return result;
}

ExcelValue eu_g39() {
  static ExcelValue result;
  if(variable_set[408] == 1) { return result;}
  result = C25;
  variable_set[408] = 1;
  return result;
}

ExcelValue eu_n39() {
  static ExcelValue result;
  if(variable_set[409] == 1) { return result;}
  result = C96;
  variable_set[409] = 1;
  return result;
}

ExcelValue eu_b40() {
  static ExcelValue result;
  if(variable_set[410] == 1) { return result;}
  result = C97;
  variable_set[410] = 1;
  return result;
}

ExcelValue eu_f40() {
  static ExcelValue result;
  if(variable_set[411] == 1) { return result;}
  result = C98;
  variable_set[411] = 1;
  return result;
}

ExcelValue eu_g40() {
  static ExcelValue result;
  if(variable_set[412] == 1) { return result;}
  result = C25;
  variable_set[412] = 1;
  return result;
}

ExcelValue eu_b41() {
  static ExcelValue result;
  if(variable_set[413] == 1) { return result;}
  result = C99;
  variable_set[413] = 1;
  return result;
}

ExcelValue eu_f41() {
  static ExcelValue result;
  if(variable_set[414] == 1) { return result;}
  result = C100;
  variable_set[414] = 1;
  return result;
}

ExcelValue eu_b42() {
  static ExcelValue result;
  if(variable_set[415] == 1) { return result;}
  result = C101;
  variable_set[415] = 1;
  return result;
}

ExcelValue eu_f42() {
  static ExcelValue result;
  if(variable_set[416] == 1) { return result;}
  result = C45;
  variable_set[416] = 1;
  return result;
}

ExcelValue eu_b43() {
  static ExcelValue result;
  if(variable_set[417] == 1) { return result;}
  result = C102;
  variable_set[417] = 1;
  return result;
}

ExcelValue eu_f43() {
  static ExcelValue result;
  if(variable_set[418] == 1) { return result;}
  result = C103;
  variable_set[418] = 1;
  return result;
}

ExcelValue eu_b45() {
  static ExcelValue result;
  if(variable_set[419] == 1) { return result;}
  result = C104;
  variable_set[419] = 1;
  return result;
}

ExcelValue eu_f45() {
  static ExcelValue result;
  if(variable_set[420] == 1) { return result;}
  result = C9;
  variable_set[420] = 1;
  return result;
}

ExcelValue eu_g45() {
  static ExcelValue result;
  if(variable_set[421] == 1) { return result;}
  result = C10;
  variable_set[421] = 1;
  return result;
}

ExcelValue eu_h45() {
  static ExcelValue result;
  if(variable_set[422] == 1) { return result;}
  result = C11;
  variable_set[422] = 1;
  return result;
}

ExcelValue eu_i45() {
  static ExcelValue result;
  if(variable_set[423] == 1) { return result;}
  result = C12;
  variable_set[423] = 1;
  return result;
}

ExcelValue eu_j45() {
  static ExcelValue result;
  if(variable_set[424] == 1) { return result;}
  result = C13;
  variable_set[424] = 1;
  return result;
}

ExcelValue eu_k45() {
  static ExcelValue result;
  if(variable_set[425] == 1) { return result;}
  result = C14;
  variable_set[425] = 1;
  return result;
}

ExcelValue eu_l45() {
  static ExcelValue result;
  if(variable_set[426] == 1) { return result;}
  result = C15;
  variable_set[426] = 1;
  return result;
}

ExcelValue eu_m45() {
  static ExcelValue result;
  if(variable_set[427] == 1) { return result;}
  result = C16;
  variable_set[427] = 1;
  return result;
}

ExcelValue eu_b46() {
  static ExcelValue result;
  if(variable_set[428] == 1) { return result;}
  result = C39;
  variable_set[428] = 1;
  return result;
}

ExcelValue eu_f46() {
  static ExcelValue result;
  if(variable_set[429] == 1) { return result;}
  result = multiply(C103,_common91());
  variable_set[429] = 1;
  return result;
}

ExcelValue eu_g46() {
  static ExcelValue result;
  if(variable_set[430] == 1) { return result;}
  result = multiply(C103,_common94());
  variable_set[430] = 1;
  return result;
}

ExcelValue eu_h46() {
  static ExcelValue result;
  if(variable_set[431] == 1) { return result;}
  result = multiply(C103,_common97());
  variable_set[431] = 1;
  return result;
}

ExcelValue eu_i46() {
  static ExcelValue result;
  if(variable_set[432] == 1) { return result;}
  result = multiply(C103,_common100());
  variable_set[432] = 1;
  return result;
}

ExcelValue eu_j46() {
  static ExcelValue result;
  if(variable_set[433] == 1) { return result;}
  result = multiply(C103,_common103());
  variable_set[433] = 1;
  return result;
}

ExcelValue eu_k46() {
  static ExcelValue result;
  if(variable_set[434] == 1) { return result;}
  result = multiply(C103,_common106());
  variable_set[434] = 1;
  return result;
}

ExcelValue eu_l46() {
  static ExcelValue result;
  if(variable_set[435] == 1) { return result;}
  result = multiply(C103,_common109());
  variable_set[435] = 1;
  return result;
}

ExcelValue eu_m46() {
  static ExcelValue result;
  if(variable_set[436] == 1) { return result;}
  result = multiply(C103,_common112());
  variable_set[436] = 1;
  return result;
}

ExcelValue eu_n46() {
  static ExcelValue result;
  if(variable_set[437] == 1) { return result;}
  result = C105;
  variable_set[437] = 1;
  return result;
}

// end EU

// start OLD UK
ExcelValue old_uk_b2() {
  static ExcelValue result;
  if(variable_set[438] == 1) { return result;}
  result = C106;
  variable_set[438] = 1;
  return result;
}

ExcelValue old_uk_b4() {
  static ExcelValue result;
  if(variable_set[439] == 1) { return result;}
  result = C107;
  variable_set[439] = 1;
  return result;
}

ExcelValue old_uk_c4() {
  static ExcelValue result;
  if(variable_set[440] == 1) { return result;}
  result = C9;
  variable_set[440] = 1;
  return result;
}

ExcelValue old_uk_d4() {
  static ExcelValue result;
  if(variable_set[441] == 1) { return result;}
  result = C10;
  variable_set[441] = 1;
  return result;
}

ExcelValue old_uk_e4() {
  static ExcelValue result;
  if(variable_set[442] == 1) { return result;}
  result = C11;
  variable_set[442] = 1;
  return result;
}

ExcelValue old_uk_f4() {
  static ExcelValue result;
  if(variable_set[443] == 1) { return result;}
  result = C12;
  variable_set[443] = 1;
  return result;
}

ExcelValue old_uk_g4() {
  static ExcelValue result;
  if(variable_set[444] == 1) { return result;}
  result = C13;
  variable_set[444] = 1;
  return result;
}

ExcelValue old_uk_h4() {
  static ExcelValue result;
  if(variable_set[445] == 1) { return result;}
  result = C14;
  variable_set[445] = 1;
  return result;
}

ExcelValue old_uk_i4() {
  static ExcelValue result;
  if(variable_set[446] == 1) { return result;}
  result = C15;
  variable_set[446] = 1;
  return result;
}

ExcelValue old_uk_j4() {
  static ExcelValue result;
  if(variable_set[447] == 1) { return result;}
  result = C16;
  variable_set[447] = 1;
  return result;
}

ExcelValue old_uk_l4() {
  static ExcelValue result;
  if(variable_set[448] == 1) { return result;}
  result = C17;
  variable_set[448] = 1;
  return result;
}

ExcelValue old_uk_b5() {
  static ExcelValue result;
  if(variable_set[449] == 1) { return result;}
  result = C18;
  variable_set[449] = 1;
  return result;
}

ExcelValue old_uk_c5() {
  static ExcelValue result;
  if(variable_set[450] == 1) { return result;}
  result = C108;
  variable_set[450] = 1;
  return result;
}

ExcelValue old_uk_d5() {
  static ExcelValue result;
  if(variable_set[451] == 1) { return result;}
  result = C109;
  variable_set[451] = 1;
  return result;
}

ExcelValue old_uk_e5() {
  static ExcelValue result;
  if(variable_set[452] == 1) { return result;}
  result = C110;
  variable_set[452] = 1;
  return result;
}

ExcelValue old_uk_f5() {
  static ExcelValue result;
  if(variable_set[453] == 1) { return result;}
  result = C111;
  variable_set[453] = 1;
  return result;
}

ExcelValue old_uk_g5() {
  static ExcelValue result;
  if(variable_set[454] == 1) { return result;}
  result = multiply(C111,C70);
  variable_set[454] = 1;
  return result;
}

ExcelValue old_uk_h5() {
  static ExcelValue result;
  if(variable_set[455] == 1) { return result;}
  result = multiply(old_uk_g5(),C70);
  variable_set[455] = 1;
  return result;
}

ExcelValue old_uk_i5() {
  static ExcelValue result;
  if(variable_set[456] == 1) { return result;}
  result = multiply(old_uk_h5(),C70);
  variable_set[456] = 1;
  return result;
}

ExcelValue old_uk_j5() {
  static ExcelValue result;
  if(variable_set[457] == 1) { return result;}
  result = multiply(old_uk_i5(),C70);
  variable_set[457] = 1;
  return result;
}

ExcelValue old_uk_k5() {
  static ExcelValue result;
  if(variable_set[458] == 1) { return result;}
  result = C25;
  variable_set[458] = 1;
  return result;
}

ExcelValue old_uk_l5() {
  static ExcelValue result;
  if(variable_set[459] == 1) { return result;}
  result = C71;
  variable_set[459] = 1;
  return result;
}

ExcelValue old_uk_b6() {
  static ExcelValue result;
  if(variable_set[460] == 1) { return result;}
  result = C50;
  variable_set[460] = 1;
  return result;
}

ExcelValue old_uk_c6() {
  static ExcelValue result;
  if(variable_set[461] == 1) { return result;}
  result = C112;
  variable_set[461] = 1;
  return result;
}

ExcelValue old_uk_d6() {
  static ExcelValue result;
  if(variable_set[462] == 1) { return result;}
  result = C113;
  variable_set[462] = 1;
  return result;
}

ExcelValue old_uk_e6() {
  static ExcelValue result;
  if(variable_set[463] == 1) { return result;}
  result = C114;
  variable_set[463] = 1;
  return result;
}

ExcelValue old_uk_f6() {
  static ExcelValue result;
  if(variable_set[464] == 1) { return result;}
  result = subtract(C115,C111);
  variable_set[464] = 1;
  return result;
}

ExcelValue old_uk_g6() {
  static ExcelValue result;
  if(variable_set[465] == 1) { return result;}
  result = subtract(old_uk_g7(),old_uk_g5());
  variable_set[465] = 1;
  return result;
}

ExcelValue old_uk_h6() {
  static ExcelValue result;
  if(variable_set[466] == 1) { return result;}
  result = subtract(old_uk_h7(),old_uk_h5());
  variable_set[466] = 1;
  return result;
}

ExcelValue old_uk_i6() {
  static ExcelValue result;
  if(variable_set[467] == 1) { return result;}
  result = subtract(old_uk_i7(),old_uk_i5());
  variable_set[467] = 1;
  return result;
}

ExcelValue old_uk_j6() {
  static ExcelValue result;
  if(variable_set[468] == 1) { return result;}
  result = subtract(old_uk_j7(),old_uk_j5());
  variable_set[468] = 1;
  return result;
}

ExcelValue old_uk_k6() {
  static ExcelValue result;
  if(variable_set[469] == 1) { return result;}
  result = C116;
  variable_set[469] = 1;
  return result;
}

ExcelValue old_uk_l6() {
  static ExcelValue result;
  if(variable_set[470] == 1) { return result;}
  result = negative(subtract(power(divide(old_uk_j6(),C112),C117),C63));
  variable_set[470] = 1;
  return result;
}

ExcelValue old_uk_b7() {
  static ExcelValue result;
  if(variable_set[471] == 1) { return result;}
  result = C39;
  variable_set[471] = 1;
  return result;
}

ExcelValue old_uk_c7() {
  static ExcelValue result;
  if(variable_set[472] == 1) { return result;}
  result = C118;
  variable_set[472] = 1;
  return result;
}

ExcelValue old_uk_d7() {
  static ExcelValue result;
  if(variable_set[473] == 1) { return result;}
  result = C119;
  variable_set[473] = 1;
  return result;
}

ExcelValue old_uk_e7() {
  static ExcelValue result;
  if(variable_set[474] == 1) { return result;}
  result = C120;
  variable_set[474] = 1;
  return result;
}

ExcelValue old_uk_f7() {
  static ExcelValue result;
  if(variable_set[475] == 1) { return result;}
  result = C115;
  variable_set[475] = 1;
  return result;
}

ExcelValue old_uk_g7() {
  static ExcelValue result;
  if(variable_set[476] == 1) { return result;}
  result = multiply(C115,C70);
  variable_set[476] = 1;
  return result;
}

ExcelValue old_uk_h7() {
  static ExcelValue result;
  if(variable_set[477] == 1) { return result;}
  result = multiply(old_uk_g7(),C70);
  variable_set[477] = 1;
  return result;
}

ExcelValue old_uk_i7() {
  static ExcelValue result;
  if(variable_set[478] == 1) { return result;}
  result = multiply(old_uk_h7(),C70);
  variable_set[478] = 1;
  return result;
}

ExcelValue old_uk_j7() {
  static ExcelValue result;
  if(variable_set[479] == 1) { return result;}
  result = multiply(old_uk_i7(),C70);
  variable_set[479] = 1;
  return result;
}

ExcelValue old_uk_k7() {
  static ExcelValue result;
  if(variable_set[480] == 1) { return result;}
  result = C121;
  variable_set[480] = 1;
  return result;
}

ExcelValue old_uk_l7() {
  static ExcelValue result;
  if(variable_set[481] == 1) { return result;}
  result = C71;
  variable_set[481] = 1;
  return result;
}

ExcelValue old_uk_b9() {
  static ExcelValue result;
  if(variable_set[482] == 1) { return result;}
  result = C122;
  variable_set[482] = 1;
  return result;
}

ExcelValue old_uk_c9() {
  static ExcelValue result;
  if(variable_set[483] == 1) { return result;}
  result = C9;
  variable_set[483] = 1;
  return result;
}

ExcelValue old_uk_d9() {
  static ExcelValue result;
  if(variable_set[484] == 1) { return result;}
  result = C10;
  variable_set[484] = 1;
  return result;
}

ExcelValue old_uk_e9() {
  static ExcelValue result;
  if(variable_set[485] == 1) { return result;}
  result = C11;
  variable_set[485] = 1;
  return result;
}

ExcelValue old_uk_f9() {
  static ExcelValue result;
  if(variable_set[486] == 1) { return result;}
  result = C12;
  variable_set[486] = 1;
  return result;
}

ExcelValue old_uk_g9() {
  static ExcelValue result;
  if(variable_set[487] == 1) { return result;}
  result = C13;
  variable_set[487] = 1;
  return result;
}

ExcelValue old_uk_h9() {
  static ExcelValue result;
  if(variable_set[488] == 1) { return result;}
  result = C14;
  variable_set[488] = 1;
  return result;
}

ExcelValue old_uk_i9() {
  static ExcelValue result;
  if(variable_set[489] == 1) { return result;}
  result = C15;
  variable_set[489] = 1;
  return result;
}

ExcelValue old_uk_j9() {
  static ExcelValue result;
  if(variable_set[490] == 1) { return result;}
  result = C16;
  variable_set[490] = 1;
  return result;
}

ExcelValue old_uk_b10() {
  static ExcelValue result;
  if(variable_set[491] == 1) { return result;}
  result = C18;
  variable_set[491] = 1;
  return result;
}

ExcelValue old_uk_c10() {
  static ExcelValue result;
  if(variable_set[492] == 1) { return result;}
  result = C63;
  variable_set[492] = 1;
  return result;
}

ExcelValue old_uk_d10() {
  static ExcelValue result;
  if(variable_set[493] == 1) { return result;}
  result = C63;
  variable_set[493] = 1;
  return result;
}

ExcelValue old_uk_e10() {
  static ExcelValue result;
  if(variable_set[494] == 1) { return result;}
  result = C63;
  variable_set[494] = 1;
  return result;
}

ExcelValue old_uk_f10() {
  static ExcelValue result;
  if(variable_set[495] == 1) { return result;}
  result = C63;
  variable_set[495] = 1;
  return result;
}

ExcelValue old_uk_g10() {
  static ExcelValue result;
  if(variable_set[496] == 1) { return result;}
  result = C63;
  variable_set[496] = 1;
  return result;
}

ExcelValue old_uk_h10() {
  static ExcelValue result;
  if(variable_set[497] == 1) { return result;}
  result = C63;
  variable_set[497] = 1;
  return result;
}

ExcelValue old_uk_i10() {
  static ExcelValue result;
  if(variable_set[498] == 1) { return result;}
  result = C63;
  variable_set[498] = 1;
  return result;
}

ExcelValue old_uk_j10() {
  static ExcelValue result;
  if(variable_set[499] == 1) { return result;}
  result = C63;
  variable_set[499] = 1;
  return result;
}

ExcelValue old_uk_k10() {
  static ExcelValue result;
  if(variable_set[500] == 1) { return result;}
  result = C44;
  variable_set[500] = 1;
  return result;
}

ExcelValue old_uk_b11() {
  static ExcelValue result;
  if(variable_set[501] == 1) { return result;}
  result = C50;
  variable_set[501] = 1;
  return result;
}

ExcelValue old_uk_c11() {
  static ExcelValue result;
  if(variable_set[502] == 1) { return result;}
  result = C123;
  variable_set[502] = 1;
  return result;
}

ExcelValue old_uk_d11() {
  static ExcelValue result;
  if(variable_set[503] == 1) { return result;}
  result = C124;
  variable_set[503] = 1;
  return result;
}

ExcelValue old_uk_e11() {
  static ExcelValue result;
  if(variable_set[504] == 1) { return result;}
  result = C125;
  variable_set[504] = 1;
  return result;
}

ExcelValue old_uk_f11() {
  static ExcelValue result;
  if(variable_set[505] == 1) { return result;}
  result = add(C125,C126);
  variable_set[505] = 1;
  return result;
}

ExcelValue old_uk_g11() {
  static ExcelValue result;
  if(variable_set[506] == 1) { return result;}
  result = add(old_uk_f11(),C126);
  variable_set[506] = 1;
  return result;
}

ExcelValue old_uk_h11() {
  static ExcelValue result;
  if(variable_set[507] == 1) { return result;}
  result = add(old_uk_g11(),C126);
  variable_set[507] = 1;
  return result;
}

ExcelValue old_uk_i11() {
  static ExcelValue result;
  if(variable_set[508] == 1) { return result;}
  result = add(old_uk_h11(),C126);
  variable_set[508] = 1;
  return result;
}

ExcelValue old_uk_j11() {
  static ExcelValue result;
  if(variable_set[509] == 1) { return result;}
  result = C63;
  variable_set[509] = 1;
  return result;
}

ExcelValue old_uk_k11() {
  static ExcelValue result;
  if(variable_set[510] == 1) { return result;}
  result = C44;
  variable_set[510] = 1;
  return result;
}

ExcelValue old_uk_l11() {
  static ExcelValue result;
  if(variable_set[511] == 1) { return result;}
  result = C126;
  variable_set[511] = 1;
  return result;
}

ExcelValue old_uk_b12() {
  static ExcelValue result;
  if(variable_set[512] == 1) { return result;}
  result = C39;
  variable_set[512] = 1;
  return result;
}

ExcelValue old_uk_c12() {
  static ExcelValue result;
  if(variable_set[513] == 1) { return result;}
  result = C127;
  variable_set[513] = 1;
  return result;
}

ExcelValue old_uk_d12() {
  static ExcelValue result;
  if(variable_set[514] == 1) { return result;}
  result = C128;
  variable_set[514] = 1;
  return result;
}

ExcelValue old_uk_e12() {
  static ExcelValue result;
  if(variable_set[515] == 1) { return result;}
  result = _common115();
  variable_set[515] = 1;
  return result;
}

ExcelValue old_uk_f12() {
  static ExcelValue result;
  if(variable_set[516] == 1) { return result;}
  result = _common118();
  variable_set[516] = 1;
  return result;
}

ExcelValue old_uk_g12() {
  static ExcelValue result;
  if(variable_set[517] == 1) { return result;}
  result = _common122();
  variable_set[517] = 1;
  return result;
}

ExcelValue old_uk_h12() {
  static ExcelValue result;
  if(variable_set[518] == 1) { return result;}
  result = _common126();
  variable_set[518] = 1;
  return result;
}

ExcelValue old_uk_i12() {
  static ExcelValue result;
  if(variable_set[519] == 1) { return result;}
  result = _common130();
  variable_set[519] = 1;
  return result;
}

ExcelValue old_uk_j12() {
  static ExcelValue result;
  if(variable_set[520] == 1) { return result;}
  result = _common134();
  variable_set[520] = 1;
  return result;
}

ExcelValue old_uk_k12() {
  static ExcelValue result;
  if(variable_set[521] == 1) { return result;}
  result = C44;
  variable_set[521] = 1;
  return result;
}

ExcelValue old_uk_c14() {
  static ExcelValue result;
  if(variable_set[522] == 1) { return result;}
  result = C9;
  variable_set[522] = 1;
  return result;
}

ExcelValue old_uk_d14() {
  static ExcelValue result;
  if(variable_set[523] == 1) { return result;}
  result = C10;
  variable_set[523] = 1;
  return result;
}

ExcelValue old_uk_e14() {
  static ExcelValue result;
  if(variable_set[524] == 1) { return result;}
  result = C11;
  variable_set[524] = 1;
  return result;
}

ExcelValue old_uk_f14() {
  static ExcelValue result;
  if(variable_set[525] == 1) { return result;}
  result = C12;
  variable_set[525] = 1;
  return result;
}

ExcelValue old_uk_g14() {
  static ExcelValue result;
  if(variable_set[526] == 1) { return result;}
  result = C13;
  variable_set[526] = 1;
  return result;
}

ExcelValue old_uk_h14() {
  static ExcelValue result;
  if(variable_set[527] == 1) { return result;}
  result = C14;
  variable_set[527] = 1;
  return result;
}

ExcelValue old_uk_i14() {
  static ExcelValue result;
  if(variable_set[528] == 1) { return result;}
  result = C15;
  variable_set[528] = 1;
  return result;
}

ExcelValue old_uk_j14() {
  static ExcelValue result;
  if(variable_set[529] == 1) { return result;}
  result = C16;
  variable_set[529] = 1;
  return result;
}

ExcelValue old_uk_b15() {
  static ExcelValue result;
  if(variable_set[530] == 1) { return result;}
  result = C129;
  variable_set[530] = 1;
  return result;
}

ExcelValue old_uk_c15() {
  static ExcelValue result;
  if(variable_set[531] == 1) { return result;}
  result = C130;
  variable_set[531] = 1;
  return result;
}

ExcelValue old_uk_d15() {
  static ExcelValue result;
  if(variable_set[532] == 1) { return result;}
  result = C130;
  variable_set[532] = 1;
  return result;
}

ExcelValue old_uk_e15() {
  static ExcelValue result;
  if(variable_set[533] == 1) { return result;}
  result = C130;
  variable_set[533] = 1;
  return result;
}

ExcelValue old_uk_f15() {
  static ExcelValue result;
  if(variable_set[534] == 1) { return result;}
  result = C130;
  variable_set[534] = 1;
  return result;
}

ExcelValue old_uk_g15() {
  static ExcelValue result;
  if(variable_set[535] == 1) { return result;}
  result = C130;
  variable_set[535] = 1;
  return result;
}

ExcelValue old_uk_h15() {
  static ExcelValue result;
  if(variable_set[536] == 1) { return result;}
  result = C130;
  variable_set[536] = 1;
  return result;
}

ExcelValue old_uk_i15() {
  static ExcelValue result;
  if(variable_set[537] == 1) { return result;}
  result = C130;
  variable_set[537] = 1;
  return result;
}

ExcelValue old_uk_j15() {
  static ExcelValue result;
  if(variable_set[538] == 1) { return result;}
  result = C130;
  variable_set[538] = 1;
  return result;
}

ExcelValue old_uk_k15() {
  static ExcelValue result;
  if(variable_set[539] == 1) { return result;}
  result = C83;
  variable_set[539] = 1;
  return result;
}

ExcelValue old_uk_b17() {
  static ExcelValue result;
  if(variable_set[540] == 1) { return result;}
  result = C131;
  variable_set[540] = 1;
  return result;
}

ExcelValue old_uk_c17() {
  static ExcelValue result;
  if(variable_set[541] == 1) { return result;}
  result = C9;
  variable_set[541] = 1;
  return result;
}

ExcelValue old_uk_d17() {
  static ExcelValue result;
  if(variable_set[542] == 1) { return result;}
  result = C10;
  variable_set[542] = 1;
  return result;
}

ExcelValue old_uk_e17() {
  static ExcelValue result;
  if(variable_set[543] == 1) { return result;}
  result = C11;
  variable_set[543] = 1;
  return result;
}

ExcelValue old_uk_f17() {
  static ExcelValue result;
  if(variable_set[544] == 1) { return result;}
  result = C12;
  variable_set[544] = 1;
  return result;
}

ExcelValue old_uk_g17() {
  static ExcelValue result;
  if(variable_set[545] == 1) { return result;}
  result = C13;
  variable_set[545] = 1;
  return result;
}

ExcelValue old_uk_h17() {
  static ExcelValue result;
  if(variable_set[546] == 1) { return result;}
  result = C14;
  variable_set[546] = 1;
  return result;
}

ExcelValue old_uk_i17() {
  static ExcelValue result;
  if(variable_set[547] == 1) { return result;}
  result = C15;
  variable_set[547] = 1;
  return result;
}

ExcelValue old_uk_j17() {
  static ExcelValue result;
  if(variable_set[548] == 1) { return result;}
  result = C16;
  variable_set[548] = 1;
  return result;
}

ExcelValue old_uk_b18() {
  static ExcelValue result;
  if(variable_set[549] == 1) { return result;}
  result = C18;
  variable_set[549] = 1;
  return result;
}

ExcelValue old_uk_c18() {
  static ExcelValue result;
  if(variable_set[550] == 1) { return result;}
  result = C132;
  variable_set[550] = 1;
  return result;
}

ExcelValue old_uk_d18() {
  static ExcelValue result;
  if(variable_set[551] == 1) { return result;}
  result = C133;
  variable_set[551] = 1;
  return result;
}

ExcelValue old_uk_e18() {
  static ExcelValue result;
  if(variable_set[552] == 1) { return result;}
  result = C134;
  variable_set[552] = 1;
  return result;
}

ExcelValue old_uk_f18() {
  static ExcelValue result;
  if(variable_set[553] == 1) { return result;}
  result = divide(multiply(_common120(),C130),C135);
  variable_set[553] = 1;
  return result;
}

ExcelValue old_uk_g18() {
  static ExcelValue result;
  if(variable_set[554] == 1) { return result;}
  result = divide(multiply(_common124(),C130),C135);
  variable_set[554] = 1;
  return result;
}

ExcelValue old_uk_h18() {
  static ExcelValue result;
  if(variable_set[555] == 1) { return result;}
  result = divide(multiply(_common128(),C130),C135);
  variable_set[555] = 1;
  return result;
}

ExcelValue old_uk_i18() {
  static ExcelValue result;
  if(variable_set[556] == 1) { return result;}
  result = divide(multiply(_common132(),C130),C135);
  variable_set[556] = 1;
  return result;
}

ExcelValue old_uk_j18() {
  static ExcelValue result;
  if(variable_set[557] == 1) { return result;}
  result = divide(multiply(_common136(),C130),C135);
  variable_set[557] = 1;
  return result;
}

ExcelValue old_uk_k18() {
  static ExcelValue result;
  if(variable_set[558] == 1) { return result;}
  result = C105;
  variable_set[558] = 1;
  return result;
}

ExcelValue old_uk_b19() {
  static ExcelValue result;
  if(variable_set[559] == 1) { return result;}
  result = C50;
  variable_set[559] = 1;
  return result;
}

ExcelValue old_uk_c19() {
  static ExcelValue result;
  if(variable_set[560] == 1) { return result;}
  result = C136;
  variable_set[560] = 1;
  return result;
}

ExcelValue old_uk_d19() {
  static ExcelValue result;
  if(variable_set[561] == 1) { return result;}
  result = C137;
  variable_set[561] = 1;
  return result;
}

ExcelValue old_uk_e19() {
  static ExcelValue result;
  if(variable_set[562] == 1) { return result;}
  result = divide(multiply(_common117(),C130),C135);
  variable_set[562] = 1;
  return result;
}

ExcelValue old_uk_f19() {
  static ExcelValue result;
  if(variable_set[563] == 1) { return result;}
  result = divide(multiply(_common121(),C130),C135);
  variable_set[563] = 1;
  return result;
}

ExcelValue old_uk_g19() {
  static ExcelValue result;
  if(variable_set[564] == 1) { return result;}
  result = divide(multiply(_common125(),C130),C135);
  variable_set[564] = 1;
  return result;
}

ExcelValue old_uk_h19() {
  static ExcelValue result;
  if(variable_set[565] == 1) { return result;}
  result = divide(multiply(_common129(),C130),C135);
  variable_set[565] = 1;
  return result;
}

ExcelValue old_uk_i19() {
  static ExcelValue result;
  if(variable_set[566] == 1) { return result;}
  result = divide(multiply(_common133(),C130),C135);
  variable_set[566] = 1;
  return result;
}

ExcelValue old_uk_j19() {
  static ExcelValue result;
  if(variable_set[567] == 1) { return result;}
  result = divide(multiply(_common137(),C130),C135);
  variable_set[567] = 1;
  return result;
}

ExcelValue old_uk_k19() {
  static ExcelValue result;
  if(variable_set[568] == 1) { return result;}
  result = C105;
  variable_set[568] = 1;
  return result;
}

ExcelValue old_uk_b20() {
  static ExcelValue result;
  if(variable_set[569] == 1) { return result;}
  result = C39;
  variable_set[569] = 1;
  return result;
}

ExcelValue old_uk_c20() {
  static ExcelValue result;
  if(variable_set[570] == 1) { return result;}
  result = C138;
  variable_set[570] = 1;
  return result;
}

ExcelValue old_uk_d20() {
  static ExcelValue result;
  if(variable_set[571] == 1) { return result;}
  result = divide(multiply(multiply(C128,C119),C130),C135);
  variable_set[571] = 1;
  return result;
}

ExcelValue old_uk_e20() {
  static ExcelValue result;
  if(variable_set[572] == 1) { return result;}
  result = divide(multiply(multiply(_common115(),C120),C130),C135);
  variable_set[572] = 1;
  return result;
}

ExcelValue old_uk_f20() {
  static ExcelValue result;
  if(variable_set[573] == 1) { return result;}
  result = divide(multiply(multiply(_common118(),C115),C130),C135);
  variable_set[573] = 1;
  return result;
}

ExcelValue old_uk_g20() {
  static ExcelValue result;
  if(variable_set[574] == 1) { return result;}
  result = divide(multiply(multiply(_common122(),old_uk_g7()),C130),C135);
  variable_set[574] = 1;
  return result;
}

ExcelValue old_uk_h20() {
  static ExcelValue result;
  if(variable_set[575] == 1) { return result;}
  result = divide(multiply(multiply(_common126(),old_uk_h7()),C130),C135);
  variable_set[575] = 1;
  return result;
}

ExcelValue old_uk_i20() {
  static ExcelValue result;
  if(variable_set[576] == 1) { return result;}
  result = divide(multiply(multiply(_common130(),old_uk_i7()),C130),C135);
  variable_set[576] = 1;
  return result;
}

ExcelValue old_uk_j20() {
  static ExcelValue result;
  if(variable_set[577] == 1) { return result;}
  result = divide(multiply(multiply(_common134(),old_uk_j7()),C130),C135);
  variable_set[577] = 1;
  return result;
}

ExcelValue old_uk_k20() {
  static ExcelValue result;
  if(variable_set[578] == 1) { return result;}
  result = C105;
  variable_set[578] = 1;
  return result;
}

ExcelValue old_uk_k22() {
  static ExcelValue result;
  if(variable_set[579] == 1) { return result;}
  result = C96;
  variable_set[579] = 1;
  return result;
}

// end OLD UK

