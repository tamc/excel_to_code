// /Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx approximately translated into C
// First we have c versions of all the excel functions that we know
#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include <math.h>

// To run the tests at the end of this file
// cc excel_to_c_runtime; ./a.out

// FIXME: Extract a header file

// I predefine an array of ExcelValues to store calculations
// Probably bad practice. At the very least, I should make it
// link to the cell reference in some way.
#define MAX_EXCEL_VALUE_HEAP_SIZE 1000000
#define MAX_MEMORY_TO_BE_FREED_HEAP_SIZE 1000000

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
static ExcelValue hlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v);
static ExcelValue hlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v, ExcelValue match_type_v);
static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error);
static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v);
static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue row_number_v);
static ExcelValue excel_isnumber(ExcelValue number);
static ExcelValue large(ExcelValue array_v, ExcelValue k_v);
static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v);
static ExcelValue left_1(ExcelValue string_v);
static ExcelValue excel_log(ExcelValue number);
static ExcelValue excel_log_2(ExcelValue number, ExcelValue base);
static ExcelValue max(int number_of_arguments, ExcelValue *arguments);
static ExcelValue min(int number_of_arguments, ExcelValue *arguments);
static ExcelValue mmult(ExcelValue a_v, ExcelValue b_v);
static ExcelValue mod(ExcelValue a_v, ExcelValue b_v);
static ExcelValue negative(ExcelValue a_v);
static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v);
static ExcelValue power(ExcelValue a_v, ExcelValue b_v);
static ExcelValue pv_3(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v);
static ExcelValue pv_4(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v, ExcelValue d_v);
static ExcelValue pv_5(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v, ExcelValue d_v, ExcelValue e_v);
static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue rank(ExcelValue number_v, ExcelValue range_v, ExcelValue order_v);
static ExcelValue rank_2(ExcelValue number_v, ExcelValue range_v);
static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue excel_int(ExcelValue number_v);
static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments);
static ExcelValue subtotal(ExcelValue type, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments);
static ExcelValue sumif(ExcelValue check_range_v, ExcelValue criteria_v, ExcelValue sum_range_v );
static ExcelValue sumif_2(ExcelValue check_range_v, ExcelValue criteria_v);
static ExcelValue sumproduct(int number_of_arguments, ExcelValue *arguments);
static ExcelValue text(ExcelValue number_v, ExcelValue format_v);
static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v);
static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v);


// My little heap for excel values
ExcelValue cells[MAX_EXCEL_VALUE_HEAP_SIZE];
int cell_counter = 0;

#define HEAPCHECK if(cell_counter >= MAX_EXCEL_VALUE_HEAP_SIZE) { printf("ExcelValue heap full. Edit MAX_EXCEL_VALUE_HEAP_SIZE in the c source code."); exit(-1); }

// My little heap for keeping pointers to memory that I need to reclaim
void *memory_that_needs_to_be_freed[MAX_MEMORY_TO_BE_FREED_HEAP_SIZE];
int memory_that_needs_to_be_freed_counter = 0;

#define MEMORY_THAT_NEEDS_TO_BE_FREED_HEAP_CHECK 

static void free_later(void *pointer) {
	memory_that_needs_to_be_freed[memory_that_needs_to_be_freed_counter] = pointer;
	memory_that_needs_to_be_freed_counter++;
	if(memory_that_needs_to_be_freed_counter >= MAX_MEMORY_TO_BE_FREED_HEAP_SIZE) { 
		printf("Memory that needs to be freed heap full. Edit MAX_MEMORY_TO_BE_FREED_HEAP_SIZE in the c source code"); 
		exit(-1);
	}
}

static void free_all_allocated_memory() {
	int i;
	for(i = 0; i < memory_that_needs_to_be_freed_counter; i++) {
		free(memory_that_needs_to_be_freed[i]);
	}
	memory_that_needs_to_be_freed_counter = 0;
}

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
	ExcelValue new_cell = cells[cell_counter];
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
	ExcelValue *pointer = malloc(sizeof(ExcelValue)*size); // Freed later
	if(pointer == 0) {
		printf("Out of memory\n");
		exit(-1);
	}
	free_later(pointer);
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
const ExcelValue NUM = {.type = ExcelError, .number = 5};

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
			 case 5: printf("NUM\n"); break;
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

static ExcelValue excel_log(ExcelValue number) {
  return excel_log_2(number, TEN);
}

static ExcelValue excel_log_2(ExcelValue number_v, ExcelValue base_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(base_v)
	NUMBER(number_v, n)
	NUMBER(base_v, b)
	CHECK_FOR_CONVERSION_ERROR

  if(n<=0) { return NUM; }
  if(b<=0) { return NUM; }

  return	new_excel_number(log(n)/log(b));
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

static ExcelValue excel_isnumber(ExcelValue potential_number) {
  if(potential_number.type == ExcelNumber) {
    return TRUE;
  } else {
    return FALSE;
  }
}

static ExcelValue excel_if(ExcelValue condition, ExcelValue true_case, ExcelValue false_case ) {
	CHECK_FOR_PASSED_ERROR(condition)
	
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

int compare_doubles (const void *a, const void *b) {
  const double *da = (const double *) a;
  const double *db = (const double *) b;

  return (*da > *db) - (*da < *db);
}

static ExcelValue large(ExcelValue range_v, ExcelValue k_v) {
  CHECK_FOR_PASSED_ERROR(range_v)
  CHECK_FOR_PASSED_ERROR(k_v)

  int k = (int) number_from(k_v);
  CHECK_FOR_CONVERSION_ERROR;

  // Check for edge case where just a single number passed
  if(range_v.type == ExcelNumber) {
    if( k == 1 ) {
      return range_v;
    } else {
      return NUM;
    }
  }

  // Otherwise grumble if not a range
  if(!range_v.type == ExcelRange) { return VALUE; }

  // Check that our k is within bounds
  if(k < 1) { return NUM; }
  int range_size = range_v.rows * range_v.columns;

  // OK this is a really naive implementation.
  // FIXME: implement the BFPRT algorithm
  double *sorted = malloc(sizeof(double)*range_size);
  int sorted_size = 0;
  ExcelValue *array_v = range_v.array;
  ExcelValue x_v;
  int i; 
  for(i = 0; i < range_size; i++ ) {
    x_v = array_v[i];
    if(x_v.type == ExcelError) { return x_v; };
    if(x_v.type == ExcelNumber) {
      sorted[sorted_size] = x_v.number;
      sorted_size++;
    }
  }
  // Check other bound
  if(k > sorted_size) { return NUM; }

  qsort(sorted, sorted_size, sizeof (double), compare_doubles);

  ExcelValue result = new_excel_number(sorted[sorted_size - k]);
  free(sorted);
  return result;
}


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
	return excel_match(lookup_value, lookup_array, ZERO);
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
	int string_must_be_freed = 0;
	switch (string_v.type) {
  	  case ExcelString:
  		string = string_v.string;
  		break;
  	  case ExcelNumber:
		  string = malloc(20); // Freed
		  if(string == 0) {
			  printf("Out of memory");
			  exit(-1);
		  }
		  string_must_be_freed = 1;
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
	
	char *left_string = malloc(number_of_characters+1); // Freed
	if(left_string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}
	free_later(left_string);
	memcpy(left_string,string,number_of_characters);
	left_string[number_of_characters] = '\0';
	if(string_must_be_freed == 1) {
		free(string);
	}
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
  ExcelValue r;
	for(i=0;i<array_size;i++) {
    switch(array[i].type) {
      case ExcelNumber:
        total += array[i].number;
        break;
      case ExcelRange:
        r = sum( array[i].rows * array[i].columns, array[i].array );
        if(r.type == ExcelError) {
          return r;
        } else {
          total += number_from(r);
        }
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

static ExcelValue mmult_error(ExcelValue a_v, ExcelValue b_v) {
  int rows = a_v.rows > b_v.rows ? a_v.rows : b_v.rows;
  int columns = a_v.columns > b_v.columns ? a_v.columns : b_v.columns;
  int i, j;

  ExcelValue *result = (ExcelValue*) new_excel_value_array(rows*columns);

  for(i=0; i<rows; i++) {
    for(j=0; j<columns; j++) {
      result[(i*columns) + j] = VALUE;
    }
  }
  return new_excel_range(result, rows, columns);
}

static ExcelValue mmult(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
  if(a_v.type != ExcelRange) { return VALUE;}
  if(b_v.type != ExcelRange) { return VALUE;}
  if(a_v.columns != b_v.rows) { return mmult_error(a_v, b_v); }
  int n = a_v.columns;
  int a_rows = a_v.rows;
  int a_columns = a_v.columns;
  int b_columns = b_v.columns;
  ExcelValue *result = (ExcelValue*) new_excel_value_array(a_rows*b_columns);
  int i, j, k;
  double sum; 
  ExcelValue *array_a = a_v.array;
  ExcelValue *array_b = b_v.array;

  ExcelValue a;
  ExcelValue b;
  
  for(i=0; i<a_rows; i++) {
    for(j=0; j<b_columns; j++) {
      sum = 0;
      for(k=0; k<n; k++) {
        a = array_a[(i*a_columns)+k];
        b = array_b[(k*b_columns)+j];
        if(a.type != ExcelNumber) { return mmult_error(a_v, b_v); }
        if(b.type != ExcelNumber) { return mmult_error(a_v, b_v); }
        sum = sum + (a.number * b.number);
      }
      result[(i*b_columns)+j] = new_excel_number(sum);
    }
  }
  return new_excel_range(result, a_rows, b_columns);
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

static ExcelValue pv_3(ExcelValue rate_v, ExcelValue nper_v, ExcelValue pmt_v) {
  return pv_4(rate_v, nper_v, pmt_v, ZERO);
}

static ExcelValue pv_4(ExcelValue rate_v, ExcelValue nper_v, ExcelValue pmt_v, ExcelValue fv_v) {
  return pv_5(rate_v, nper_v, pmt_v, fv_v, ZERO);
}

static ExcelValue pv_5(ExcelValue rate_v, ExcelValue nper_v, ExcelValue pmt_v, ExcelValue fv_v, ExcelValue type_v ) {
  CHECK_FOR_PASSED_ERROR(rate_v)
  CHECK_FOR_PASSED_ERROR(nper_v)
  CHECK_FOR_PASSED_ERROR(pmt_v)
  CHECK_FOR_PASSED_ERROR(fv_v)
  CHECK_FOR_PASSED_ERROR(type_v)

  NUMBER(rate_v, rate)
  NUMBER(nper_v, nper)
  NUMBER(pmt_v, payment)
  NUMBER(fv_v, fv)
  NUMBER(type_v, start_of_period)
  CHECK_FOR_CONVERSION_ERROR

  if(rate< 0) {
    return VALUE;
  }

  double present_value = 0;

  // Sum up the payments
  if(rate == 0) {
    present_value = -payment * nper;
  } else {
    present_value = -payment * ((1-pow(1+rate,-nper))/rate);
  }

  // Adjust for beginning or end of period
  if(start_of_period == 0) {
   // Do Nothing
  } else if(start_of_period == 1) {
   present_value = present_value * (1+rate);
  } else {
   return VALUE;
  } 

  // Add on the final value
  present_value = present_value - (fv/pow(1+rate,nper));
  
  return new_excel_number(present_value);
}


static ExcelValue power(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
		
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
  double result = pow(a,b);
  if(isnan(result) == 1) {
    return NUM;
  } else {
    return new_excel_number(result);
  }
}
static ExcelValue rank(ExcelValue number_v, ExcelValue range_v, ExcelValue order_v) {
  CHECK_FOR_PASSED_ERROR(number_v)
  CHECK_FOR_PASSED_ERROR(range_v)
  CHECK_FOR_PASSED_ERROR(order_v)

  NUMBER(number_v, number)
  NUMBER(order_v, order)

  ExcelValue *array;
  int size;

	CHECK_FOR_CONVERSION_ERROR

  if(range_v.type != ExcelRange) {
    array = new_excel_value_array(1);
    array[0] = range_v;
    size = 1;
  } else {
    array = range_v.array;
    size = range_v.rows * range_v.columns;
  }

  int ranked = 1;
  int found = false;

  int i;
  ExcelValue cell;

  for(i=0; i<size; i++) {
    cell = array[i];
    if(cell.type == ExcelError) { return cell; }
    if(cell.type == ExcelNumber) {
      if(cell.number == number) { found = true; }
      if(order == 0) { if(cell.number > number) { ranked++; } }
      if(order != 0) { if(cell.number < number) { ranked++; } }
    }
  }
  if(found == false) { return NA; }
  return new_excel_number(ranked);
}

static ExcelValue rank_2(ExcelValue number_v, ExcelValue range_v) {
  return rank(number_v, range_v, ZERO);
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

static ExcelValue excel_int(ExcelValue number_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
		
	NUMBER(number_v, number)
	CHECK_FOR_CONVERSION_ERROR
		
	return new_excel_number(floor(number));
}

static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments) {
	int allocated_length = 100;
	int used_length = 0;
	char *string = malloc(allocated_length); // Freed later
	if(string == 0) {
	  printf("Out of memory");
	  exit(-1);
	}
	free_later(string);
	char *current_string;
	int current_string_length;
	int must_free_current_string;
	ExcelValue current_v;
	int i;
	for(i=0;i<number_of_arguments;i++) {
		must_free_current_string = 0;
		current_v = (ExcelValue) arguments[i];
		switch (current_v.type) {
  	  case ExcelString:
	  		current_string = current_v.string;
	  		break;
  	  case ExcelNumber:
			  current_string = malloc(20); // Freed
		  	if(current_string == 0) {
		  	  printf("Out of memory");
		  	  exit(-1);
		  	}
			must_free_current_string = 1;				  
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
		if(must_free_current_string == 1) {
			free(current_string);
		}
		used_length = used_length + current_string_length;
	}
	string = realloc(string,used_length+1);
  string[used_length] = '\0';
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
  ExcelComparison *criteria =  malloc(sizeof(ExcelComparison)*number_of_criteria); // freed at end of function
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
  // Tidy up
  free(criteria);
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
  ExcelValue **ranges = malloc(sizeof(ExcelValue *)*number_of_arguments); // Added free statements
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
		free(ranges);
        return current_value;
        break;
      case ExcelEmpty:
		free(ranges);
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
	free(ranges);
  	return new_excel_number(sum);
}

static ExcelValue text(ExcelValue number_v, ExcelValue format_v) {
  CHECK_FOR_PASSED_ERROR(number_v)
  CHECK_FOR_PASSED_ERROR(format_v)

	char *s;
	char *p;
	double n;
  ExcelValue result;

  if(number_v.type == ExcelEmpty) {
    number_v = ZERO;
  }

  if(format_v.type == ExcelEmpty) {
    return new_excel_string("");
  }

  if(number_v.type == ExcelString) {
 	 	s = number_v.string;
		if (s == NULL || *s == '\0' || isspace(*s)) {
			number_v = ZERO;
		}	        
		n = strtod (s, &p);
		if(*p == '\0') {
		  number_v = new_excel_number(n);
		}
  }

  if(number_v.type != ExcelNumber) {
    return number_v;
  }

  if(format_v.type != ExcelString) {
    return format_v;
  }

  if(strcmp(format_v.string,"0%") == 0) {
    // FIXME: Too little? 
    s = malloc(100);
    free_later(s);
    sprintf(s, "%d%%",(int) round(number_v.number*100));
    result = new_excel_string(s);
  } else {
    return format_v;
  }

  // inspect_excel_value(result);
  return result;
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
  if(match_type_v.type == ExcelNumber && match_type_v.number >= 0 && match_type_v.number <= 1) {
    match_type_v.type = ExcelBoolean;
  }
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

static ExcelValue hlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v) {
  return hlookup(lookup_value_v,lookup_table_v,row_number_v,TRUE);
}

static ExcelValue hlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v, ExcelValue match_type_v) {
  CHECK_FOR_PASSED_ERROR(lookup_value_v)
  CHECK_FOR_PASSED_ERROR(lookup_table_v)
  CHECK_FOR_PASSED_ERROR(row_number_v)
  CHECK_FOR_PASSED_ERROR(match_type_v)

  if(lookup_value_v.type == ExcelEmpty) return NA;
  if(lookup_table_v.type != ExcelRange) return NA;
  if(row_number_v.type != ExcelNumber) return NA;
  if(match_type_v.type == ExcelNumber && match_type_v.number >= 0 && match_type_v.number <= 1) {
    match_type_v.type = ExcelBoolean;
  }
  if(match_type_v.type != ExcelBoolean) return NA;
    
  int i;
  int last_good_match = 0;
  int rows = lookup_table_v.rows;
  int columns = lookup_table_v.columns;
  ExcelValue *array = lookup_table_v.array;
  ExcelValue possible_match_v;
  
  if(row_number_v.number > rows) return REF;
  if(row_number_v.number < 1) return VALUE;
  
  if(match_type_v.number == false) { // Exact match required
    for(i=0; i< columns; i++) {
      possible_match_v = array[i];
      if(excel_equal(lookup_value_v,possible_match_v).number == true) {
        return array[((((int) row_number_v.number)-1)*columns)+(i)];
      }
    }
    return NA;
  } else { // Highest value that is less than or equal
    for(i=0; i< columns; i++) {
      possible_match_v = array[i];
      if(lookup_value_v.type != possible_match_v.type) continue;
      if(more_than(possible_match_v,lookup_value_v).number == true) {
        if(i == 0) return NA;
        return array[((((int) row_number_v.number)-1)*columns)+(i-1)];
      } else {
        last_good_match = i;
      }
    }
    return array[((((int) row_number_v.number)-1)*columns)+(last_good_match)];
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

  // Test Large
  ExcelValue large_test_array_1[] = { new_excel_number(10), new_excel_number(100), new_excel_number(500), BLANK };
  ExcelValue large_test_array_1_v = new_excel_range(large_test_array_1, 1, 4);
  assert(large(large_test_array_1_v, new_excel_number(1)).number == 500);
  assert(large(large_test_array_1_v, new_excel_number(2)).number == 100);
  assert(large(large_test_array_1_v, new_excel_number(3)).number == 10);
  assert(large(large_test_array_1_v, new_excel_number(4)).type == ExcelError);
  assert(large(new_excel_number(500), new_excel_number(1)).number == 500);
  ExcelValue large_test_array_2[] = { new_excel_number(10), new_excel_number(100), new_excel_number(500), VALUE };
  ExcelValue large_test_array_2_v = new_excel_range(large_test_array_2, 1, 4);
  assert(large(large_test_array_2_v,new_excel_number(2)).type == ExcelError);
  assert(large(new_excel_number(500),VALUE).type == ExcelError);

	
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
	assert(excel_if(TRUE,new_excel_number(10),NA).type == ExcelNumber);
	assert(excel_if(TRUE,new_excel_number(10),NA).number == 10);
	
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
    // ... should return power of its arguments
	assert(power(new_excel_number(2),new_excel_number(3)).number == 8);
	assert(power(new_excel_number(4.0),new_excel_number(0.5)).number == 2.0);
	assert(power(new_excel_number(-4.0),new_excel_number(0.5)).type == ExcelError);
	
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

	// Test int
    assert(excel_int(new_excel_number(8.9)).number == 8.0);
    assert(excel_int(new_excel_number(-8.9)).number == -9.0);

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
  assert(string_join(2, string_join_array_1).string[11] == '\0');
	// ... should cope with an arbitrary number of arguments
  assert(string_join(3, string_join_array_2).string[11] == '!');
  assert(string_join(3, string_join_array_3).string[12] == '\0');
	// ... should convert values to strings as it goes
  assert(string_join(2, string_join_array_3).string[4] == '1');
  assert(string_join(2, string_join_array_3).string[5] == '0');
  assert(string_join(2, string_join_array_3).string[6] == '\0');
	// ... should convert integer values into strings without decimal points, and float values with decimal points
  assert(string_join(2, string_join_array_4).string[4] == '1');
  assert(string_join(2, string_join_array_4).string[5] == '0');
  assert(string_join(2, string_join_array_4).string[6] == '.');
  assert(string_join(2, string_join_array_4).string[7] == '5');
  assert(string_join(2, string_join_array_4).string[8] == '\0');
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
  // .. the four argument variant should accept 0 and 1 instead of TRUE and FALSE
  assert(vlookup(new_excel_string("HELLO"),vlookup_a2_v,new_excel_number(2),ZERO).number == 10);
  assert(vlookup(new_excel_string("HELMP"),vlookup_a2_v,new_excel_number(2),ONE).number == 10);
  // ... BLANK should not match with anything" do
  assert(vlookup_3(BLANK,vlookup_a3_v,new_excel_number(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(vlookup(VALUE,vlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),VALUE,new_excel_number(2),FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(vlookup(new_excel_number(2.0),vlookup_a1_v,new_excel_number(2),VALUE).type == ExcelError);
  assert(vlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);
	
  // Test HLOOKUP
  ExcelValue hlookup_a1[] = {new_excel_number(1),new_excel_number(2),new_excel_number(3),new_excel_number(10),new_excel_number(20),new_excel_number(30)};
  ExcelValue hlookup_a2[] = {new_excel_string("hello"),new_excel_number(2),new_excel_number(3),new_excel_number(10),new_excel_number(20),new_excel_number(30)};
  ExcelValue hlookup_a3[] = {BLANK,new_excel_number(2),new_excel_number(3),new_excel_number(10),new_excel_number(20),new_excel_number(30)};
  ExcelValue hlookup_a1_v = new_excel_range(hlookup_a1,2,3);
  ExcelValue hlookup_a2_v = new_excel_range(hlookup_a2,2,3);
  ExcelValue hlookup_a3_v = new_excel_range(hlookup_a3,2,3);
  // ... should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument
  assert(hlookup_3(new_excel_number(2.0),hlookup_a1_v,new_excel_number(2)).number == 20);
  assert(hlookup_3(new_excel_number(1.5),hlookup_a1_v,new_excel_number(2)).number == 10);
  assert(hlookup_3(new_excel_number(0.5),hlookup_a1_v,new_excel_number(2)).type == ExcelError);
  assert(hlookup_3(new_excel_number(10),hlookup_a1_v,new_excel_number(2)).number == 30);
  assert(hlookup_3(new_excel_number(2.6),hlookup_a1_v,new_excel_number(2)).number == 20);
  // ... has a four argument variant that matches the lookup type
  assert(hlookup(new_excel_number(2.6),hlookup_a1_v,new_excel_number(2),TRUE).number == 20);
  assert(hlookup(new_excel_number(2.6),hlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(hlookup(new_excel_string("HELLO"),hlookup_a2_v,new_excel_number(2),FALSE).number == 10);
  assert(hlookup(new_excel_string("HELMP"),hlookup_a2_v,new_excel_number(2),TRUE).number == 10);
  // ... that four argument variant should accept 0 or 1 for the lookup type
  assert(hlookup(new_excel_number(2.6),hlookup_a1_v,new_excel_number(2),ONE).number == 20);
  assert(hlookup(new_excel_number(2.6),hlookup_a1_v,new_excel_number(2),ZERO).type == ExcelError);
  assert(hlookup(new_excel_string("HELLO"),hlookup_a2_v,new_excel_number(2),ZERO).number == 10);
  assert(hlookup(new_excel_string("HELMP"),hlookup_a2_v,new_excel_number(2),ONE).number == 10);
  // ... BLANK should not match with anything" do
  assert(hlookup_3(BLANK,hlookup_a3_v,new_excel_number(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(hlookup(VALUE,hlookup_a1_v,new_excel_number(2),FALSE).type == ExcelError);
  assert(hlookup(new_excel_number(2.0),VALUE,new_excel_number(2),FALSE).type == ExcelError);
  assert(hlookup(new_excel_number(2.0),hlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(hlookup(new_excel_number(2.0),hlookup_a1_v,new_excel_number(2),VALUE).type == ExcelError);
  assert(hlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);

  // Test SUM
  ExcelValue sum_array_0[] = {new_excel_number(1084.4557258064517),new_excel_number(32.0516914516129),new_excel_number(137.36439193548387)};
  ExcelValue sum_array_0_v = new_excel_range(sum_array_0,3,1);
  ExcelValue sum_array_1[] = {sum_array_0_v};
  assert(sum(1,sum_array_1).number == 1253.8718091935484);

  // Test PV
  assert((int) pv_3(new_excel_number(0.03), new_excel_number(12), new_excel_number(100)).number == -995);
  assert((int) pv_4(new_excel_number(0.03), new_excel_number(12), new_excel_number(-100), new_excel_number(100)).number == 925);
  assert((int) pv_5(new_excel_number(0.03), new_excel_number(12), new_excel_number(-100), new_excel_number(-100), new_excel_number(1)).number == 1095);
  
  // Test TEXT
  assert(strcmp(text(new_excel_number(1.0), new_excel_string("0%")).string, "100%") == 0);
  assert(strcmp(text(new_excel_string("1"), new_excel_string("0%")).string, "100%") == 0);
  assert(strcmp(text(BLANK, new_excel_string("0%")).string, "0%") == 0);
  assert(strcmp(text(new_excel_number(1.0), BLANK).string, "") == 0);
  assert(strcmp(text(new_excel_string("ASGASD"), new_excel_string("0%")).string, "ASGASD") == 0);

  // Test LOG
  // One argument variant assumes LOG base 10
	assert(excel_log(new_excel_number(10)).number == 1);
	assert(excel_log(new_excel_number(100)).number == 2);
	assert(excel_log(new_excel_number(0)).type == ExcelError);
  // Two argument variant allows LOG base to be specified
	assert(excel_log_2(new_excel_number(8),new_excel_number(2)).number == 3.0);
	assert(excel_log_2(new_excel_number(8),new_excel_number(0)).type == ExcelError);
  
  // Test MMULT (Matrix multiplication)
  ExcelValue mmult_1[] = { ONE, TWO, THREE, FOUR};
  ExcelValue mmult_2[] = { FOUR, THREE, TWO, ONE};
  ExcelValue mmult_3[] = { ONE, TWO};
  ExcelValue mmult_4[] = { THREE, FOUR};
  ExcelValue mmult_5[] = { ONE, BLANK, THREE, FOUR};

  ExcelValue mmult_1_v = new_excel_range( mmult_1, 2, 2);
  ExcelValue mmult_2_v = new_excel_range( mmult_2, 2, 2);
  ExcelValue mmult_3_v = new_excel_range( mmult_3, 1, 2);
  ExcelValue mmult_4_v = new_excel_range( mmult_4, 2, 1);
  ExcelValue mmult_5_v = new_excel_range( mmult_5, 2, 2);

  // Treat the ranges as matrices and multiply them
  ExcelValue mmult_result_1_v = mmult(mmult_1_v, mmult_2_v);
  assert(mmult_result_1_v.type == ExcelRange);
  assert(mmult_result_1_v.rows == 2);
  assert(mmult_result_1_v.columns == 2);
  ExcelValue *mmult_result_1_a = mmult_result_1_v.array;
  assert(mmult_result_1_a[0].number == 8);
  assert(mmult_result_1_a[1].number == 5);
  assert(mmult_result_1_a[2].number == 20);
  assert(mmult_result_1_a[3].number == 13);
  
  ExcelValue mmult_result_2_v = mmult(mmult_3_v, mmult_4_v);
  assert(mmult_result_2_v.type == ExcelRange);
  assert(mmult_result_2_v.rows == 1);
  assert(mmult_result_2_v.columns == 1);
  ExcelValue *mmult_result_2_a = mmult_result_2_v.array;
  assert(mmult_result_2_a[0].number == 11);

  // Return an error if any cells are not numbers
  ExcelValue mmult_result_3_v = mmult(mmult_5_v, mmult_2_v);
  assert(mmult_result_3_v.type == ExcelRange);
  assert(mmult_result_3_v.rows == 2);
  assert(mmult_result_3_v.columns == 2);
  ExcelValue *mmult_result_3_a = mmult_result_3_v.array;
  assert(mmult_result_3_a[0].type == ExcelError);
  assert(mmult_result_3_a[1].type == ExcelError);
  assert(mmult_result_3_a[2].type == ExcelError);
  assert(mmult_result_3_a[3].type == ExcelError);
  
  // Returns errors if arguments are not ranges
  // FIXME: Should work in edge case where passed two numbers which excel treats as ranges with one row and column
  ExcelValue mmult_result_4_v = mmult(ONE, mmult_2_v);
  assert(mmult_result_4_v.type == ExcelError);
  
  // Returns errors if the ranges aren't the right size to multiply
  ExcelValue mmult_result_5_v = mmult(mmult_1_v, mmult_3_v);
  assert(mmult_result_5_v.type == ExcelRange);
  assert(mmult_result_5_v.rows == 2);
  assert(mmult_result_5_v.columns == 2);
  ExcelValue *mmult_result_5_a = mmult_result_5_v.array;
  assert(mmult_result_5_a[0].type == ExcelError);
  assert(mmult_result_5_a[1].type == ExcelError);
  assert(mmult_result_5_a[2].type == ExcelError);
  assert(mmult_result_5_a[3].type == ExcelError);

  // Test the RANK() function
  ExcelValue rank_1_a[] = { FIVE, BLANK, THREE, ONE, ONE, FOUR, FIVE, TRUE, SIX, new_excel_string("Hi")};
  ExcelValue rank_2_a[] = { FIVE, BLANK, THREE, NA, ONE, FOUR, FIVE, TRUE, SIX, new_excel_string("Hi")};
  ExcelValue rank_1_v = new_excel_range( rank_1_a, 2, 5);
  ExcelValue rank_2_v = new_excel_range( rank_2_a, 2, 5);

  // Basics
  assert(rank(THREE, rank_1_v, ZERO).number == 5);
  assert(rank_2(THREE, rank_1_v).number == 5);
  assert(rank(THREE, rank_1_v, ONE).number == 3);
  assert(rank(ONE, rank_1_v, ZERO).number == 6);
  assert(rank(new_excel_string("3"), rank_1_v, ONE).number == 3);

  // Errors
  assert(rank(TEN, rank_1_v, ZERO).type == ExcelError);
  assert(rank(THREE, rank_2_v, ZERO).type == ExcelError);


  // Test the ISNUMBER function
  assert(excel_isnumber(ONE).type == ExcelBoolean);
  assert(excel_isnumber(ONE).number == 1);
  assert(excel_isnumber(BLANK).type == ExcelBoolean);
  assert(excel_isnumber(BLANK).number == 0);
  assert(excel_isnumber(new_excel_string("Hello")).type == ExcelBoolean);
  assert(excel_isnumber(new_excel_string("Hello")).number == 0);
  assert(excel_isnumber(TRUE).type == ExcelBoolean);
  assert(excel_isnumber(TRUE).number == 0);

  // Release memory
  free_all_allocated_memory();
  
  return 0;
}

int main() {
	return test_functions();
}
// End of the generic c functions

// Start of the file specific functions

// definitions
static ExcelValue ORIGINAL_EXCEL_FILENAME = {.type = ExcelString, .string = "/Users/tamc/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx" };
ExcelValue valuetypes_a1();
ExcelValue valuetypes_a2();
ExcelValue valuetypes_a3();
ExcelValue valuetypes_a4();
ExcelValue valuetypes_a5();
ExcelValue valuetypes_a6();
ExcelValue formulaetypes_a1();
ExcelValue formulaetypes_b1();
ExcelValue formulaetypes_a2();
ExcelValue formulaetypes_b2();
ExcelValue formulaetypes_a3();
ExcelValue formulaetypes_b3();
ExcelValue formulaetypes_a4();
ExcelValue formulaetypes_b4();
ExcelValue formulaetypes_a5();
ExcelValue formulaetypes_b5();
ExcelValue formulaetypes_a6();
ExcelValue formulaetypes_b6();
ExcelValue formulaetypes_a7();
ExcelValue formulaetypes_b7();
ExcelValue formulaetypes_a8();
ExcelValue formulaetypes_b8();
ExcelValue ranges_b1();
ExcelValue ranges_c1();
ExcelValue ranges_a2();
ExcelValue ranges_b2();
ExcelValue ranges_c2();
ExcelValue ranges_a3();
ExcelValue ranges_b3();
ExcelValue ranges_c3();
ExcelValue ranges_a4();
ExcelValue ranges_b4();
ExcelValue ranges_c4();
ExcelValue ranges_f4();
ExcelValue ranges_e5();
ExcelValue ranges_f5();
ExcelValue ranges_g5();
ExcelValue ranges_f6();
ExcelValue referencing_a1();
ExcelValue referencing_a2();
ExcelValue referencing_a4();
ExcelValue referencing_b4();
ExcelValue referencing_c4();
ExcelValue referencing_a5();
ExcelValue referencing_b8();
ExcelValue referencing_b9();
ExcelValue referencing_b11();
ExcelValue referencing_c11();
ExcelValue referencing_c15();
ExcelValue referencing_d15();
ExcelValue referencing_e15();
ExcelValue referencing_f15();
ExcelValue referencing_c16();
ExcelValue referencing_d16();
ExcelValue referencing_e16();
ExcelValue referencing_f16();
ExcelValue referencing_c17();
ExcelValue referencing_d17();
ExcelValue referencing_e17();
ExcelValue referencing_f17();
ExcelValue referencing_c18();
ExcelValue referencing_d18();
ExcelValue referencing_e18();
ExcelValue referencing_f18();
ExcelValue referencing_c19();
ExcelValue referencing_d19();
ExcelValue referencing_e19();
ExcelValue referencing_f19();
ExcelValue referencing_c22();
ExcelValue referencing_d22();
ExcelValue referencing_d23();
ExcelValue referencing_d24();
ExcelValue referencing_d25();
ExcelValue referencing_c31();
ExcelValue referencing_o31();
ExcelValue referencing_f33();
ExcelValue referencing_g33();
ExcelValue referencing_h33();
ExcelValue referencing_i33();
ExcelValue referencing_j33();
ExcelValue referencing_k33();
ExcelValue referencing_l33();
ExcelValue referencing_m33();
ExcelValue referencing_n33();
ExcelValue referencing_o33();
ExcelValue referencing_c34();
ExcelValue referencing_d34();
ExcelValue referencing_e34();
ExcelValue referencing_f34();
ExcelValue referencing_g34();
ExcelValue referencing_h34();
ExcelValue referencing_i34();
ExcelValue referencing_j34();
ExcelValue referencing_k34();
ExcelValue referencing_l34();
ExcelValue referencing_m34();
ExcelValue referencing_n34();
ExcelValue referencing_c35();
ExcelValue referencing_d35();
ExcelValue referencing_j35();
ExcelValue referencing_m35();
ExcelValue referencing_n35();
ExcelValue referencing_o35();
ExcelValue referencing_c36();
ExcelValue referencing_d36();
ExcelValue referencing_j36();
ExcelValue referencing_m36();
ExcelValue referencing_n36();
ExcelValue referencing_o36();
ExcelValue referencing_c37();
ExcelValue referencing_d37();
ExcelValue referencing_f37();
ExcelValue referencing_m37();
ExcelValue referencing_n37();
ExcelValue referencing_o37();
ExcelValue referencing_c38();
ExcelValue referencing_d38();
ExcelValue referencing_i38();
ExcelValue referencing_m38();
ExcelValue referencing_n38();
ExcelValue referencing_o38();
ExcelValue referencing_c39();
ExcelValue referencing_d39();
ExcelValue referencing_e39();
ExcelValue referencing_h39();
ExcelValue referencing_m39();
ExcelValue referencing_n39();
ExcelValue referencing_o39();
ExcelValue referencing_c40();
ExcelValue referencing_d40();
ExcelValue referencing_e40();
ExcelValue referencing_g40();
ExcelValue referencing_j40();
ExcelValue referencing_m40();
ExcelValue referencing_n40();
ExcelValue referencing_o40();
ExcelValue referencing_c41();
ExcelValue referencing_d41();
ExcelValue referencing_e41();
ExcelValue referencing_g41();
ExcelValue referencing_j41();
ExcelValue referencing_m41();
ExcelValue referencing_n41();
ExcelValue referencing_o41();
ExcelValue referencing_c42();
ExcelValue referencing_d42();
ExcelValue referencing_f42();
ExcelValue referencing_l42();
ExcelValue referencing_m42();
ExcelValue referencing_o42();
ExcelValue referencing_c43();
ExcelValue referencing_d43();
ExcelValue referencing_f43();
ExcelValue referencing_l43();
ExcelValue referencing_m43();
ExcelValue referencing_o43();
ExcelValue referencing_c44();
ExcelValue referencing_d44();
ExcelValue referencing_l44();
ExcelValue referencing_m44();
ExcelValue referencing_n44();
ExcelValue referencing_o44();
ExcelValue referencing_c45();
ExcelValue referencing_d45();
ExcelValue referencing_g45();
ExcelValue referencing_j45();
ExcelValue referencing_m45();
ExcelValue referencing_n45();
ExcelValue referencing_o45();
ExcelValue referencing_c46();
ExcelValue referencing_d46();
ExcelValue referencing_g46();
ExcelValue referencing_h46();
ExcelValue referencing_m46();
ExcelValue referencing_n46();
ExcelValue referencing_o46();
ExcelValue referencing_c47();
ExcelValue referencing_d47();
ExcelValue referencing_e47();
ExcelValue referencing_k47();
ExcelValue referencing_m47();
ExcelValue referencing_n47();
ExcelValue referencing_o47();
ExcelValue referencing_d50();
ExcelValue referencing_g50();
ExcelValue referencing_d51();
ExcelValue referencing_g51();
ExcelValue referencing_d52();
ExcelValue referencing_g52();
ExcelValue referencing_d53();
ExcelValue referencing_g53();
ExcelValue referencing_d54();
ExcelValue referencing_g54();
ExcelValue referencing_d55();
ExcelValue referencing_g55();
ExcelValue referencing_d56();
ExcelValue referencing_g56();
ExcelValue referencing_d57();
ExcelValue referencing_g57();
ExcelValue referencing_d58();
ExcelValue referencing_g58();
ExcelValue referencing_d59();
ExcelValue referencing_g59();
ExcelValue referencing_d60();
ExcelValue referencing_g60();
ExcelValue referencing_d61();
ExcelValue referencing_g61();
ExcelValue referencing_d62();
ExcelValue referencing_g62();
ExcelValue referencing_d64();
ExcelValue referencing_e64();
ExcelValue referencing_h64();
ExcelValue tables_a1();
ExcelValue tables_b2();
ExcelValue tables_c2();
ExcelValue tables_d2();
ExcelValue tables_b3();
ExcelValue tables_c3();
ExcelValue tables_d3();
ExcelValue tables_b4();
ExcelValue tables_c4();
ExcelValue tables_d4();
ExcelValue tables_f4();
ExcelValue tables_g4();
ExcelValue tables_h4();
ExcelValue tables_b5();
ExcelValue tables_c5();
ExcelValue tables_e6();
ExcelValue tables_f6();
ExcelValue tables_g6();
ExcelValue tables_e7();
ExcelValue tables_f7();
ExcelValue tables_g7();
ExcelValue tables_e8();
ExcelValue tables_f8();
ExcelValue tables_g8();
ExcelValue tables_e9();
ExcelValue tables_f9();
ExcelValue tables_g9();
ExcelValue tables_c10();
ExcelValue tables_e10();
ExcelValue tables_f10();
ExcelValue tables_g10();
ExcelValue tables_c11();
ExcelValue tables_e11();
ExcelValue tables_f11();
ExcelValue tables_g11();
ExcelValue tables_c12();
ExcelValue tables_c13();
ExcelValue tables_c14();
ExcelValue s_innapropriate_sheet_name__c4();
ExcelValue _common0();
ExcelValue _common1();
ExcelValue _common3();
ExcelValue _common7();
ExcelValue _common9();
// end of definitions

// Used to decide whether to recalculate a cell
static int variable_set[256];

// Used to reset all cached values and free up memory
void reset() {
  int i;
  cell_counter = 0;
  free_all_allocated_memory(); 
  for(i = 0; i < 256; i++) {
    variable_set[i] = 0;
  }
};

// starting the value constants
static ExcelValue constant1 = {.type = ExcelString, .string = "Hello"};
static ExcelValue constant2 = {.type = ExcelNumber, .number = 1.0};
static ExcelValue constant3 = {.type = ExcelNumber, .number = 3.1415};
static ExcelValue constant4 = {.type = ExcelString, .string = "Simple"};
static ExcelValue constant5 = {.type = ExcelNumber, .number = 2.0};
static ExcelValue constant6 = {.type = ExcelString, .string = "Sharing"};
static ExcelValue constant7 = {.type = ExcelNumber, .number = 267.7467614837482};
static ExcelValue constant8 = {.type = ExcelString, .string = "Shared"};
static ExcelValue constant9 = {.type = ExcelString, .string = "Array (single)"};
static ExcelValue constant10 = {.type = ExcelString, .string = "Arraying (multiple)"};
static ExcelValue constant11 = {.type = ExcelString, .string = "Not Eight"};
static ExcelValue constant12 = {.type = ExcelString, .string = "Arrayed (multiple)"};
static ExcelValue constant13 = {.type = ExcelString, .string = "This sheet"};
static ExcelValue constant14 = {.type = ExcelString, .string = "Other sheet"};
static ExcelValue constant15 = {.type = ExcelString, .string = "Standard"};
static ExcelValue constant16 = {.type = ExcelString, .string = "Column"};
static ExcelValue constant17 = {.type = ExcelString, .string = "Row"};
static ExcelValue constant18 = {.type = ExcelNumber, .number = 3.0};
static ExcelValue constant19 = {.type = ExcelNumber, .number = 10.0};
static ExcelValue constant20 = {.type = ExcelString, .string = "Named"};
static ExcelValue constant21 = {.type = ExcelString, .string = "Reference"};
static ExcelValue constant22 = {.type = ExcelNumber, .number = 4.0};
static ExcelValue constant23 = {.type = ExcelNumber, .number = 1.4535833325868115};
static ExcelValue constant24 = {.type = ExcelNumber, .number = 1.511726665890284};
static ExcelValue constant25 = {.type = ExcelNumber, .number = 1.5407983325420203};
static ExcelValue constant26 = {.type = ExcelNumber, .number = 9.054545454545455};
static ExcelValue constant27 = {.type = ExcelNumber, .number = 12.0};
static ExcelValue constant28 = {.type = ExcelNumber, .number = 18.0};
static ExcelValue constant29 = {.type = ExcelNumber, .number = 0.3681150635671386};
static ExcelValue constant30 = {.type = ExcelNumber, .number = 0.40588480110308967};
static ExcelValue constant31 = {.type = ExcelNumber, .number = 0.42190146532760275};
static ExcelValue constant32 = {.type = ExcelNumber, .number = 0.651};
static ExcelValue constant33 = {.type = ExcelString, .string = "Technology efficiencies -- hot water -- annual mean"};
static ExcelValue constant34 = {.type = ExcelString, .string = "% of input energy"};
static ExcelValue constant35 = {.type = ExcelString, .string = "Electricity (delivered to end user)"};
static ExcelValue constant36 = {.type = ExcelString, .string = "Electricity (supplied to grid)"};
static ExcelValue constant37 = {.type = ExcelString, .string = "Solid hydrocarbons"};
static ExcelValue constant38 = {.type = ExcelString, .string = "Liquid hydrocarbons"};
static ExcelValue constant39 = {.type = ExcelString, .string = "Gaseous hydrocarbons"};
static ExcelValue constant40 = {.type = ExcelString, .string = "Heat transport"};
static ExcelValue constant41 = {.type = ExcelString, .string = "Environmental heat"};
static ExcelValue constant42 = {.type = ExcelString, .string = "Heating & cooling"};
static ExcelValue constant43 = {.type = ExcelString, .string = "Conversion losses"};
static ExcelValue constant44 = {.type = ExcelString, .string = "Balance"};
static ExcelValue constant45 = {.type = ExcelString, .string = "Code"};
static ExcelValue constant46 = {.type = ExcelString, .string = "Technology"};
static ExcelValue constant47 = {.type = ExcelString, .string = "Notes"};
static ExcelValue constant48 = {.type = ExcelString, .string = "V.01"};
static ExcelValue constant49 = {.type = ExcelString, .string = "V.02"};
static ExcelValue constant50 = {.type = ExcelString, .string = "V.03"};
static ExcelValue constant51 = {.type = ExcelString, .string = "V.04"};
static ExcelValue constant52 = {.type = ExcelString, .string = "V.05"};
static ExcelValue constant53 = {.type = ExcelString, .string = "V.07"};
static ExcelValue constant54 = {.type = ExcelString, .string = "R.07"};
static ExcelValue constant55 = {.type = ExcelString, .string = "H.01"};
static ExcelValue constant56 = {.type = ExcelString, .string = "X.01"};
static ExcelValue constant57 = {.type = ExcelString, .string = "Gas boiler (old)"};
static ExcelValue constant58 = {.type = ExcelNumber, .number = -1.0};
static ExcelValue constant59 = {.type = ExcelNumber, .number = 0.76};
static ExcelValue constant60 = {.type = ExcelNumber, .number = 0.24};
static ExcelValue constant61 = {.type = ExcelNumber, .number = 0.0};
static ExcelValue constant62 = {.type = ExcelString, .string = "Gas boiler (new)"};
static ExcelValue constant63 = {.type = ExcelNumber, .number = 0.91};
static ExcelValue constant64 = {.type = ExcelNumber, .number = 0.09};
static ExcelValue constant65 = {.type = ExcelString, .string = "Resistive heating"};
static ExcelValue constant66 = {.type = ExcelString, .string = "Oil-fired boiler"};
static ExcelValue constant67 = {.type = ExcelNumber, .number = 0.97};
static ExcelValue constant68 = {.type = ExcelNumber, .number = 0.03};
static ExcelValue constant69 = {.type = ExcelNumber, .number = -2.7755575615628914e-17};
static ExcelValue constant70 = {.type = ExcelNumber, .number = 5.0};
static ExcelValue constant71 = {.type = ExcelString, .string = "Solid-fuel boiler"};
static ExcelValue constant72 = {.type = ExcelString, .string = "[2]"};
static ExcelValue constant73 = {.type = ExcelNumber, .number = 0.87};
static ExcelValue constant74 = {.type = ExcelNumber, .number = 0.13};
static ExcelValue constant75 = {.type = ExcelNumber, .number = 6.0};
static ExcelValue constant76 = {.type = ExcelString, .string = "Stirling engine micro-CHP"};
static ExcelValue constant77 = {.type = ExcelString, .string = "[3]"};
static ExcelValue constant78 = {.type = ExcelNumber, .number = 0.225};
static ExcelValue constant79 = {.type = ExcelNumber, .number = 0.63};
static ExcelValue constant80 = {.type = ExcelNumber, .number = 0.145};
static ExcelValue constant81 = {.type = ExcelNumber, .number = 7.0};
static ExcelValue constant82 = {.type = ExcelString, .string = "Fuel-cell micro-CHP"};
static ExcelValue constant83 = {.type = ExcelNumber, .number = 0.45};
static ExcelValue constant84 = {.type = ExcelNumber, .number = 0.1};
static ExcelValue constant85 = {.type = ExcelNumber, .number = 8.0};
static ExcelValue constant86 = {.type = ExcelString, .string = "Air-source heat pump"};
static ExcelValue constant87 = {.type = ExcelNumber, .number = 9.0};
static ExcelValue constant88 = {.type = ExcelString, .string = "Ground-source heat pump"};
static ExcelValue constant89 = {.type = ExcelNumber, .number = -2.0};
static ExcelValue constant90 = {.type = ExcelString, .string = "Geothermal electricity"};
static ExcelValue constant91 = {.type = ExcelNumber, .number = 0.85};
static ExcelValue constant92 = {.type = ExcelNumber, .number = 0.15};
static ExcelValue constant93 = {.type = ExcelNumber, .number = 11.0};
static ExcelValue constant94 = {.type = ExcelString, .string = "Community scale gas CHP with local district heating"};
static ExcelValue constant95 = {.type = ExcelNumber, .number = 0.38};
static ExcelValue constant96 = {.type = ExcelString, .string = "Community scale solid-fuel CHP with local district heating"};
static ExcelValue constant97 = {.type = ExcelNumber, .number = 0.17};
static ExcelValue constant98 = {.type = ExcelNumber, .number = 0.57};
static ExcelValue constant99 = {.type = ExcelNumber, .number = 0.26};
static ExcelValue constant100 = {.type = ExcelNumber, .number = 13.0};
static ExcelValue constant101 = {.type = ExcelString, .string = "Long distance district heating from large power stations"};
static ExcelValue constant102 = {.type = ExcelString, .string = "[6]"};
static ExcelValue constant103 = {.type = ExcelNumber, .number = 0.9};
static ExcelValue constant104 = {.type = ExcelNumber, .number = 137.26515207025273};
static ExcelValue constant105 = {.type = ExcelNumber, .number = 30.731004194832696};
static ExcelValue constant106 = {.type = ExcelNumber, .number = 20.487336129888465};
static ExcelValue constant107 = {.type = ExcelNumber, .number = 8.194934451955387};
static ExcelValue constant108 = {.type = ExcelNumber, .number = 0};
static ExcelValue constant109 = {.type = ExcelString, .string = "ColA"};
static ExcelValue constant110 = {.type = ExcelString, .string = "ColB"};
static ExcelValue constant111 = {.type = ExcelString, .string = "Column1"};
static ExcelValue constant112 = {.type = ExcelString, .string = "A"};
static ExcelValue constant113 = {.type = ExcelString, .string = "B"};
static ExcelValue constant114 = {.type = ExcelString, .string = "2B"};
// ending the value constants

ExcelValue valuetypes_a1_default() {
  return TRUE;
}
static ExcelValue valuetypes_a1_variable;
ExcelValue valuetypes_a1() { if(variable_set[0] == 1) { return valuetypes_a1_variable; } else { return valuetypes_a1_default(); } }
void set_valuetypes_a1(ExcelValue newValue) { variable_set[0] = 1; valuetypes_a1_variable = newValue; }

ExcelValue valuetypes_a2_default() {
  return constant1;
}
static ExcelValue valuetypes_a2_variable;
ExcelValue valuetypes_a2() { if(variable_set[1] == 1) { return valuetypes_a2_variable; } else { return valuetypes_a2_default(); } }
void set_valuetypes_a2(ExcelValue newValue) { variable_set[1] = 1; valuetypes_a2_variable = newValue; }

ExcelValue valuetypes_a3_default() {
  return constant2;
}
static ExcelValue valuetypes_a3_variable;
ExcelValue valuetypes_a3() { if(variable_set[2] == 1) { return valuetypes_a3_variable; } else { return valuetypes_a3_default(); } }
void set_valuetypes_a3(ExcelValue newValue) { variable_set[2] = 1; valuetypes_a3_variable = newValue; }

ExcelValue valuetypes_a4_default() {
  return constant3;
}
static ExcelValue valuetypes_a4_variable;
ExcelValue valuetypes_a4() { if(variable_set[3] == 1) { return valuetypes_a4_variable; } else { return valuetypes_a4_default(); } }
void set_valuetypes_a4(ExcelValue newValue) { variable_set[3] = 1; valuetypes_a4_variable = newValue; }

ExcelValue valuetypes_a5_default() {
  return NAME;
}
static ExcelValue valuetypes_a5_variable;
ExcelValue valuetypes_a5() { if(variable_set[4] == 1) { return valuetypes_a5_variable; } else { return valuetypes_a5_default(); } }
void set_valuetypes_a5(ExcelValue newValue) { variable_set[4] = 1; valuetypes_a5_variable = newValue; }

ExcelValue valuetypes_a6_default() {
  return constant1;
}
static ExcelValue valuetypes_a6_variable;
ExcelValue valuetypes_a6() { if(variable_set[5] == 1) { return valuetypes_a6_variable; } else { return valuetypes_a6_default(); } }
void set_valuetypes_a6(ExcelValue newValue) { variable_set[5] = 1; valuetypes_a6_variable = newValue; }

ExcelValue formulaetypes_a1() {
  static ExcelValue result;
  if(variable_set[6] == 1) { return result;}
  result = constant4;
  variable_set[6] = 1;
  return result;
}

ExcelValue formulaetypes_b1() {
  static ExcelValue result;
  if(variable_set[7] == 1) { return result;}
  result = constant5;
  variable_set[7] = 1;
  return result;
}

ExcelValue formulaetypes_a2() {
  static ExcelValue result;
  if(variable_set[8] == 1) { return result;}
  result = constant6;
  variable_set[8] = 1;
  return result;
}

ExcelValue formulaetypes_b2() {
  static ExcelValue result;
  if(variable_set[9] == 1) { return result;}
  result = constant7;
  variable_set[9] = 1;
  return result;
}

ExcelValue formulaetypes_a3() {
  static ExcelValue result;
  if(variable_set[10] == 1) { return result;}
  result = constant8;
  variable_set[10] = 1;
  return result;
}

ExcelValue formulaetypes_b3() {
  static ExcelValue result;
  if(variable_set[11] == 1) { return result;}
  result = constant7;
  variable_set[11] = 1;
  return result;
}

ExcelValue formulaetypes_a4() {
  static ExcelValue result;
  if(variable_set[12] == 1) { return result;}
  result = constant8;
  variable_set[12] = 1;
  return result;
}

ExcelValue formulaetypes_b4() {
  static ExcelValue result;
  if(variable_set[13] == 1) { return result;}
  result = constant7;
  variable_set[13] = 1;
  return result;
}

ExcelValue formulaetypes_a5() {
  static ExcelValue result;
  if(variable_set[14] == 1) { return result;}
  result = constant9;
  variable_set[14] = 1;
  return result;
}

ExcelValue formulaetypes_b5() {
  static ExcelValue result;
  if(variable_set[15] == 1) { return result;}
  result = constant5;
  variable_set[15] = 1;
  return result;
}

ExcelValue formulaetypes_a6() {
  static ExcelValue result;
  if(variable_set[16] == 1) { return result;}
  result = constant10;
  variable_set[16] = 1;
  return result;
}

ExcelValue formulaetypes_b6() {
  static ExcelValue result;
  if(variable_set[17] == 1) { return result;}
  result = constant11;
  variable_set[17] = 1;
  return result;
}

ExcelValue formulaetypes_a7() {
  static ExcelValue result;
  if(variable_set[18] == 1) { return result;}
  result = constant12;
  variable_set[18] = 1;
  return result;
}

ExcelValue formulaetypes_b7() {
  static ExcelValue result;
  if(variable_set[19] == 1) { return result;}
  result = constant11;
  variable_set[19] = 1;
  return result;
}

ExcelValue formulaetypes_a8() {
  static ExcelValue result;
  if(variable_set[20] == 1) { return result;}
  result = constant12;
  variable_set[20] = 1;
  return result;
}

ExcelValue formulaetypes_b8() {
  static ExcelValue result;
  if(variable_set[21] == 1) { return result;}
  result = constant11;
  variable_set[21] = 1;
  return result;
}

ExcelValue ranges_b1() {
  static ExcelValue result;
  if(variable_set[22] == 1) { return result;}
  result = constant13;
  variable_set[22] = 1;
  return result;
}

ExcelValue ranges_c1() {
  static ExcelValue result;
  if(variable_set[23] == 1) { return result;}
  result = constant14;
  variable_set[23] = 1;
  return result;
}

ExcelValue ranges_a2() {
  static ExcelValue result;
  if(variable_set[24] == 1) { return result;}
  result = constant15;
  variable_set[24] = 1;
  return result;
}

ExcelValue ranges_b2() {
  static ExcelValue result;
  if(variable_set[25] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = ranges_f4();
  array1[1] = ranges_f5();
  array1[2] = ranges_f6();
  ExcelValue array1_ev = new_excel_range(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[25] = 1;
  return result;
}

ExcelValue ranges_c2() {
  static ExcelValue result;
  if(variable_set[26] == 1) { return result;}
  static ExcelValue array1[2];
  array1[0] = valuetypes_a3();
  array1[1] = valuetypes_a4();
  ExcelValue array1_ev = new_excel_range(array1,2,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[26] = 1;
  return result;
}

ExcelValue ranges_a3() {
  static ExcelValue result;
  if(variable_set[27] == 1) { return result;}
  result = constant16;
  variable_set[27] = 1;
  return result;
}

ExcelValue ranges_b3() {
  static ExcelValue result;
  if(variable_set[28] == 1) { return result;}
  static ExcelValue array1[6];
  array1[0] = BLANK;
  array1[1] = BLANK;
  array1[2] = BLANK;
  array1[3] = ranges_f4();
  array1[4] = ranges_f5();
  array1[5] = ranges_f6();
  ExcelValue array1_ev = new_excel_range(array1,6,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[28] = 1;
  return result;
}

ExcelValue ranges_c3() {
  static ExcelValue result;
  if(variable_set[29] == 1) { return result;}
  static ExcelValue array1[6];
  array1[0] = valuetypes_a1();
  array1[1] = valuetypes_a2();
  array1[2] = valuetypes_a3();
  array1[3] = valuetypes_a4();
  array1[4] = valuetypes_a5();
  array1[5] = valuetypes_a6();
  ExcelValue array1_ev = new_excel_range(array1,6,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[29] = 1;
  return result;
}

ExcelValue ranges_a4() {
  static ExcelValue result;
  if(variable_set[30] == 1) { return result;}
  result = constant17;
  variable_set[30] = 1;
  return result;
}

ExcelValue ranges_b4() {
  static ExcelValue result;
  if(variable_set[31] == 1) { return result;}
  static ExcelValue array1[7];
  array1[0] = BLANK;
  array1[1] = BLANK;
  array1[2] = BLANK;
  array1[3] = BLANK;
  array1[4] = ranges_e5();
  array1[5] = ranges_f5();
  array1[6] = ranges_g5();
  ExcelValue array1_ev = new_excel_range(array1,1,7);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[31] = 1;
  return result;
}

ExcelValue ranges_c4() {
  static ExcelValue result;
  if(variable_set[32] == 1) { return result;}
  ExcelValue array0[] = {valuetypes_a4()};
  result = sum(1, array0);
  variable_set[32] = 1;
  return result;
}

ExcelValue ranges_f4_default() {
  return constant2;
}
static ExcelValue ranges_f4_variable;
ExcelValue ranges_f4() { if(variable_set[33] == 1) { return ranges_f4_variable; } else { return ranges_f4_default(); } }
void set_ranges_f4(ExcelValue newValue) { variable_set[33] = 1; ranges_f4_variable = newValue; }

ExcelValue ranges_e5_default() {
  return constant2;
}
static ExcelValue ranges_e5_variable;
ExcelValue ranges_e5() { if(variable_set[34] == 1) { return ranges_e5_variable; } else { return ranges_e5_default(); } }
void set_ranges_e5(ExcelValue newValue) { variable_set[34] = 1; ranges_e5_variable = newValue; }

ExcelValue ranges_f5_default() {
  return constant5;
}
static ExcelValue ranges_f5_variable;
ExcelValue ranges_f5() { if(variable_set[35] == 1) { return ranges_f5_variable; } else { return ranges_f5_default(); } }
void set_ranges_f5(ExcelValue newValue) { variable_set[35] = 1; ranges_f5_variable = newValue; }

ExcelValue ranges_g5_default() {
  return constant18;
}
static ExcelValue ranges_g5_variable;
ExcelValue ranges_g5() { if(variable_set[36] == 1) { return ranges_g5_variable; } else { return ranges_g5_default(); } }
void set_ranges_g5(ExcelValue newValue) { variable_set[36] = 1; ranges_g5_variable = newValue; }

ExcelValue ranges_f6_default() {
  return constant18;
}
static ExcelValue ranges_f6_variable;
ExcelValue ranges_f6() { if(variable_set[37] == 1) { return ranges_f6_variable; } else { return ranges_f6_default(); } }
void set_ranges_f6(ExcelValue newValue) { variable_set[37] = 1; ranges_f6_variable = newValue; }

ExcelValue referencing_a1() {
  static ExcelValue result;
  if(variable_set[38] == 1) { return result;}
  result = referencing_c4();
  variable_set[38] = 1;
  return result;
}

ExcelValue referencing_a2() {
  static ExcelValue result;
  if(variable_set[39] == 1) { return result;}
  result = referencing_c4();
  variable_set[39] = 1;
  return result;
}

ExcelValue referencing_a4_default() {
  return constant19;
}
static ExcelValue referencing_a4_variable;
ExcelValue referencing_a4() { if(variable_set[40] == 1) { return referencing_a4_variable; } else { return referencing_a4_default(); } }
void set_referencing_a4(ExcelValue newValue) { variable_set[40] = 1; referencing_a4_variable = newValue; }

ExcelValue referencing_b4() {
  static ExcelValue result;
  if(variable_set[41] == 1) { return result;}
  result = _common0();
  variable_set[41] = 1;
  return result;
}

ExcelValue referencing_c4() {
  static ExcelValue result;
  if(variable_set[42] == 1) { return result;}
  result = add(_common0(),constant2);
  variable_set[42] = 1;
  return result;
}

ExcelValue referencing_a5() {
  static ExcelValue result;
  if(variable_set[43] == 1) { return result;}
  result = constant18;
  variable_set[43] = 1;
  return result;
}

ExcelValue referencing_b8() {
  static ExcelValue result;
  if(variable_set[44] == 1) { return result;}
  result = referencing_c4();
  variable_set[44] = 1;
  return result;
}

ExcelValue referencing_b9() {
  static ExcelValue result;
  if(variable_set[45] == 1) { return result;}
  result = _common1();
  variable_set[45] = 1;
  return result;
}

ExcelValue referencing_b11() {
  static ExcelValue result;
  if(variable_set[46] == 1) { return result;}
  result = constant20;
  variable_set[46] = 1;
  return result;
}

ExcelValue referencing_c11() {
  static ExcelValue result;
  if(variable_set[47] == 1) { return result;}
  result = constant21;
  variable_set[47] = 1;
  return result;
}

ExcelValue referencing_c15_default() {
  return constant2;
}
static ExcelValue referencing_c15_variable;
ExcelValue referencing_c15() { if(variable_set[48] == 1) { return referencing_c15_variable; } else { return referencing_c15_default(); } }
void set_referencing_c15(ExcelValue newValue) { variable_set[48] = 1; referencing_c15_variable = newValue; }

ExcelValue referencing_d15_default() {
  return constant5;
}
static ExcelValue referencing_d15_variable;
ExcelValue referencing_d15() { if(variable_set[49] == 1) { return referencing_d15_variable; } else { return referencing_d15_default(); } }
void set_referencing_d15(ExcelValue newValue) { variable_set[49] = 1; referencing_d15_variable = newValue; }

ExcelValue referencing_e15_default() {
  return constant18;
}
static ExcelValue referencing_e15_variable;
ExcelValue referencing_e15() { if(variable_set[50] == 1) { return referencing_e15_variable; } else { return referencing_e15_default(); } }
void set_referencing_e15(ExcelValue newValue) { variable_set[50] = 1; referencing_e15_variable = newValue; }

ExcelValue referencing_f15_default() {
  return constant22;
}
static ExcelValue referencing_f15_variable;
ExcelValue referencing_f15() { if(variable_set[51] == 1) { return referencing_f15_variable; } else { return referencing_f15_default(); } }
void set_referencing_f15(ExcelValue newValue) { variable_set[51] = 1; referencing_f15_variable = newValue; }

ExcelValue referencing_c16_default() {
  return constant23;
}
static ExcelValue referencing_c16_variable;
ExcelValue referencing_c16() { if(variable_set[52] == 1) { return referencing_c16_variable; } else { return referencing_c16_default(); } }
void set_referencing_c16(ExcelValue newValue) { variable_set[52] = 1; referencing_c16_variable = newValue; }

ExcelValue referencing_d16_default() {
  return constant23;
}
static ExcelValue referencing_d16_variable;
ExcelValue referencing_d16() { if(variable_set[53] == 1) { return referencing_d16_variable; } else { return referencing_d16_default(); } }
void set_referencing_d16(ExcelValue newValue) { variable_set[53] = 1; referencing_d16_variable = newValue; }

ExcelValue referencing_e16_default() {
  return constant24;
}
static ExcelValue referencing_e16_variable;
ExcelValue referencing_e16() { if(variable_set[54] == 1) { return referencing_e16_variable; } else { return referencing_e16_default(); } }
void set_referencing_e16(ExcelValue newValue) { variable_set[54] = 1; referencing_e16_variable = newValue; }

ExcelValue referencing_f16_default() {
  return constant25;
}
static ExcelValue referencing_f16_variable;
ExcelValue referencing_f16() { if(variable_set[55] == 1) { return referencing_f16_variable; } else { return referencing_f16_default(); } }
void set_referencing_f16(ExcelValue newValue) { variable_set[55] = 1; referencing_f16_variable = newValue; }

ExcelValue referencing_c17_default() {
  return constant26;
}
static ExcelValue referencing_c17_variable;
ExcelValue referencing_c17() { if(variable_set[56] == 1) { return referencing_c17_variable; } else { return referencing_c17_default(); } }
void set_referencing_c17(ExcelValue newValue) { variable_set[56] = 1; referencing_c17_variable = newValue; }

ExcelValue referencing_d17_default() {
  return constant27;
}
static ExcelValue referencing_d17_variable;
ExcelValue referencing_d17() { if(variable_set[57] == 1) { return referencing_d17_variable; } else { return referencing_d17_default(); } }
void set_referencing_d17(ExcelValue newValue) { variable_set[57] = 1; referencing_d17_variable = newValue; }

ExcelValue referencing_e17_default() {
  return constant28;
}
static ExcelValue referencing_e17_variable;
ExcelValue referencing_e17() { if(variable_set[58] == 1) { return referencing_e17_variable; } else { return referencing_e17_default(); } }
void set_referencing_e17(ExcelValue newValue) { variable_set[58] = 1; referencing_e17_variable = newValue; }

ExcelValue referencing_f17_default() {
  return constant28;
}
static ExcelValue referencing_f17_variable;
ExcelValue referencing_f17() { if(variable_set[59] == 1) { return referencing_f17_variable; } else { return referencing_f17_default(); } }
void set_referencing_f17(ExcelValue newValue) { variable_set[59] = 1; referencing_f17_variable = newValue; }

ExcelValue referencing_c18_default() {
  return constant29;
}
static ExcelValue referencing_c18_variable;
ExcelValue referencing_c18() { if(variable_set[60] == 1) { return referencing_c18_variable; } else { return referencing_c18_default(); } }
void set_referencing_c18(ExcelValue newValue) { variable_set[60] = 1; referencing_c18_variable = newValue; }

ExcelValue referencing_d18_default() {
  return constant29;
}
static ExcelValue referencing_d18_variable;
ExcelValue referencing_d18() { if(variable_set[61] == 1) { return referencing_d18_variable; } else { return referencing_d18_default(); } }
void set_referencing_d18(ExcelValue newValue) { variable_set[61] = 1; referencing_d18_variable = newValue; }

ExcelValue referencing_e18_default() {
  return constant30;
}
static ExcelValue referencing_e18_variable;
ExcelValue referencing_e18() { if(variable_set[62] == 1) { return referencing_e18_variable; } else { return referencing_e18_default(); } }
void set_referencing_e18(ExcelValue newValue) { variable_set[62] = 1; referencing_e18_variable = newValue; }

ExcelValue referencing_f18_default() {
  return constant31;
}
static ExcelValue referencing_f18_variable;
ExcelValue referencing_f18() { if(variable_set[63] == 1) { return referencing_f18_variable; } else { return referencing_f18_default(); } }
void set_referencing_f18(ExcelValue newValue) { variable_set[63] = 1; referencing_f18_variable = newValue; }

ExcelValue referencing_c19_default() {
  return constant32;
}
static ExcelValue referencing_c19_variable;
ExcelValue referencing_c19() { if(variable_set[64] == 1) { return referencing_c19_variable; } else { return referencing_c19_default(); } }
void set_referencing_c19(ExcelValue newValue) { variable_set[64] = 1; referencing_c19_variable = newValue; }

ExcelValue referencing_d19_default() {
  return constant32;
}
static ExcelValue referencing_d19_variable;
ExcelValue referencing_d19() { if(variable_set[65] == 1) { return referencing_d19_variable; } else { return referencing_d19_default(); } }
void set_referencing_d19(ExcelValue newValue) { variable_set[65] = 1; referencing_d19_variable = newValue; }

ExcelValue referencing_e19_default() {
  return constant32;
}
static ExcelValue referencing_e19_variable;
ExcelValue referencing_e19() { if(variable_set[66] == 1) { return referencing_e19_variable; } else { return referencing_e19_default(); } }
void set_referencing_e19(ExcelValue newValue) { variable_set[66] = 1; referencing_e19_variable = newValue; }

ExcelValue referencing_f19_default() {
  return constant32;
}
static ExcelValue referencing_f19_variable;
ExcelValue referencing_f19() { if(variable_set[67] == 1) { return referencing_f19_variable; } else { return referencing_f19_default(); } }
void set_referencing_f19(ExcelValue newValue) { variable_set[67] = 1; referencing_f19_variable = newValue; }

ExcelValue referencing_c22_default() {
  return constant22;
}
static ExcelValue referencing_c22_variable;
ExcelValue referencing_c22() { if(variable_set[68] == 1) { return referencing_c22_variable; } else { return referencing_c22_default(); } }
void set_referencing_c22(ExcelValue newValue) { variable_set[68] = 1; referencing_c22_variable = newValue; }

ExcelValue referencing_d22() {
  static ExcelValue result;
  if(variable_set[69] == 1) { return result;}
  result = excel_index(_common3(),constant2,constant2);
  variable_set[69] = 1;
  return result;
}

ExcelValue referencing_d23() {
  static ExcelValue result;
  if(variable_set[70] == 1) { return result;}
  result = excel_index(_common3(),constant5,constant2);
  variable_set[70] = 1;
  return result;
}

ExcelValue referencing_d24() {
  static ExcelValue result;
  if(variable_set[71] == 1) { return result;}
  result = excel_index(_common3(),constant18,constant2);
  variable_set[71] = 1;
  return result;
}

ExcelValue referencing_d25() {
  static ExcelValue result;
  if(variable_set[72] == 1) { return result;}
  result = excel_index(_common3(),constant22,constant2);
  variable_set[72] = 1;
  return result;
}

ExcelValue referencing_c31() {
  static ExcelValue result;
  if(variable_set[73] == 1) { return result;}
  result = constant33;
  variable_set[73] = 1;
  return result;
}

ExcelValue referencing_o31() {
  static ExcelValue result;
  if(variable_set[74] == 1) { return result;}
  result = constant34;
  variable_set[74] = 1;
  return result;
}

ExcelValue referencing_f33() {
  static ExcelValue result;
  if(variable_set[75] == 1) { return result;}
  result = constant35;
  variable_set[75] = 1;
  return result;
}

ExcelValue referencing_g33() {
  static ExcelValue result;
  if(variable_set[76] == 1) { return result;}
  result = constant36;
  variable_set[76] = 1;
  return result;
}

ExcelValue referencing_h33() {
  static ExcelValue result;
  if(variable_set[77] == 1) { return result;}
  result = constant37;
  variable_set[77] = 1;
  return result;
}

ExcelValue referencing_i33() {
  static ExcelValue result;
  if(variable_set[78] == 1) { return result;}
  result = constant38;
  variable_set[78] = 1;
  return result;
}

ExcelValue referencing_j33() {
  static ExcelValue result;
  if(variable_set[79] == 1) { return result;}
  result = constant39;
  variable_set[79] = 1;
  return result;
}

ExcelValue referencing_k33() {
  static ExcelValue result;
  if(variable_set[80] == 1) { return result;}
  result = constant40;
  variable_set[80] = 1;
  return result;
}

ExcelValue referencing_l33() {
  static ExcelValue result;
  if(variable_set[81] == 1) { return result;}
  result = constant41;
  variable_set[81] = 1;
  return result;
}

ExcelValue referencing_m33() {
  static ExcelValue result;
  if(variable_set[82] == 1) { return result;}
  result = constant42;
  variable_set[82] = 1;
  return result;
}

ExcelValue referencing_n33() {
  static ExcelValue result;
  if(variable_set[83] == 1) { return result;}
  result = constant43;
  variable_set[83] = 1;
  return result;
}

ExcelValue referencing_o33() {
  static ExcelValue result;
  if(variable_set[84] == 1) { return result;}
  result = constant44;
  variable_set[84] = 1;
  return result;
}

ExcelValue referencing_c34() {
  static ExcelValue result;
  if(variable_set[85] == 1) { return result;}
  result = constant45;
  variable_set[85] = 1;
  return result;
}

ExcelValue referencing_d34() {
  static ExcelValue result;
  if(variable_set[86] == 1) { return result;}
  result = constant46;
  variable_set[86] = 1;
  return result;
}

ExcelValue referencing_e34() {
  static ExcelValue result;
  if(variable_set[87] == 1) { return result;}
  result = constant47;
  variable_set[87] = 1;
  return result;
}

ExcelValue referencing_f34_default() {
  return constant48;
}
static ExcelValue referencing_f34_variable;
ExcelValue referencing_f34() { if(variable_set[88] == 1) { return referencing_f34_variable; } else { return referencing_f34_default(); } }
void set_referencing_f34(ExcelValue newValue) { variable_set[88] = 1; referencing_f34_variable = newValue; }

ExcelValue referencing_g34_default() {
  return constant49;
}
static ExcelValue referencing_g34_variable;
ExcelValue referencing_g34() { if(variable_set[89] == 1) { return referencing_g34_variable; } else { return referencing_g34_default(); } }
void set_referencing_g34(ExcelValue newValue) { variable_set[89] = 1; referencing_g34_variable = newValue; }

ExcelValue referencing_h34_default() {
  return constant50;
}
static ExcelValue referencing_h34_variable;
ExcelValue referencing_h34() { if(variable_set[90] == 1) { return referencing_h34_variable; } else { return referencing_h34_default(); } }
void set_referencing_h34(ExcelValue newValue) { variable_set[90] = 1; referencing_h34_variable = newValue; }

ExcelValue referencing_i34_default() {
  return constant51;
}
static ExcelValue referencing_i34_variable;
ExcelValue referencing_i34() { if(variable_set[91] == 1) { return referencing_i34_variable; } else { return referencing_i34_default(); } }
void set_referencing_i34(ExcelValue newValue) { variable_set[91] = 1; referencing_i34_variable = newValue; }

ExcelValue referencing_j34_default() {
  return constant52;
}
static ExcelValue referencing_j34_variable;
ExcelValue referencing_j34() { if(variable_set[92] == 1) { return referencing_j34_variable; } else { return referencing_j34_default(); } }
void set_referencing_j34(ExcelValue newValue) { variable_set[92] = 1; referencing_j34_variable = newValue; }

ExcelValue referencing_k34_default() {
  return constant53;
}
static ExcelValue referencing_k34_variable;
ExcelValue referencing_k34() { if(variable_set[93] == 1) { return referencing_k34_variable; } else { return referencing_k34_default(); } }
void set_referencing_k34(ExcelValue newValue) { variable_set[93] = 1; referencing_k34_variable = newValue; }

ExcelValue referencing_l34_default() {
  return constant54;
}
static ExcelValue referencing_l34_variable;
ExcelValue referencing_l34() { if(variable_set[94] == 1) { return referencing_l34_variable; } else { return referencing_l34_default(); } }
void set_referencing_l34(ExcelValue newValue) { variable_set[94] = 1; referencing_l34_variable = newValue; }

ExcelValue referencing_m34_default() {
  return constant55;
}
static ExcelValue referencing_m34_variable;
ExcelValue referencing_m34() { if(variable_set[95] == 1) { return referencing_m34_variable; } else { return referencing_m34_default(); } }
void set_referencing_m34(ExcelValue newValue) { variable_set[95] = 1; referencing_m34_variable = newValue; }

ExcelValue referencing_n34_default() {
  return constant56;
}
static ExcelValue referencing_n34_variable;
ExcelValue referencing_n34() { if(variable_set[96] == 1) { return referencing_n34_variable; } else { return referencing_n34_default(); } }
void set_referencing_n34(ExcelValue newValue) { variable_set[96] = 1; referencing_n34_variable = newValue; }

ExcelValue referencing_c35() {
  static ExcelValue result;
  if(variable_set[97] == 1) { return result;}
  result = constant2;
  variable_set[97] = 1;
  return result;
}

ExcelValue referencing_d35() {
  static ExcelValue result;
  if(variable_set[98] == 1) { return result;}
  result = constant57;
  variable_set[98] = 1;
  return result;
}

ExcelValue referencing_j35_default() {
  return constant58;
}
static ExcelValue referencing_j35_variable;
ExcelValue referencing_j35() { if(variable_set[99] == 1) { return referencing_j35_variable; } else { return referencing_j35_default(); } }
void set_referencing_j35(ExcelValue newValue) { variable_set[99] = 1; referencing_j35_variable = newValue; }

ExcelValue referencing_m35_default() {
  return constant59;
}
static ExcelValue referencing_m35_variable;
ExcelValue referencing_m35() { if(variable_set[100] == 1) { return referencing_m35_variable; } else { return referencing_m35_default(); } }
void set_referencing_m35(ExcelValue newValue) { variable_set[100] = 1; referencing_m35_variable = newValue; }

ExcelValue referencing_n35_default() {
  return constant60;
}
static ExcelValue referencing_n35_variable;
ExcelValue referencing_n35() { if(variable_set[101] == 1) { return referencing_n35_variable; } else { return referencing_n35_default(); } }
void set_referencing_n35(ExcelValue newValue) { variable_set[101] = 1; referencing_n35_variable = newValue; }

ExcelValue referencing_o35() {
  static ExcelValue result;
  if(variable_set[102] == 1) { return result;}
  result = constant61;
  variable_set[102] = 1;
  return result;
}

ExcelValue referencing_c36() {
  static ExcelValue result;
  if(variable_set[103] == 1) { return result;}
  result = constant5;
  variable_set[103] = 1;
  return result;
}

ExcelValue referencing_d36() {
  static ExcelValue result;
  if(variable_set[104] == 1) { return result;}
  result = constant62;
  variable_set[104] = 1;
  return result;
}

ExcelValue referencing_j36_default() {
  return constant58;
}
static ExcelValue referencing_j36_variable;
ExcelValue referencing_j36() { if(variable_set[105] == 1) { return referencing_j36_variable; } else { return referencing_j36_default(); } }
void set_referencing_j36(ExcelValue newValue) { variable_set[105] = 1; referencing_j36_variable = newValue; }

ExcelValue referencing_m36_default() {
  return constant63;
}
static ExcelValue referencing_m36_variable;
ExcelValue referencing_m36() { if(variable_set[106] == 1) { return referencing_m36_variable; } else { return referencing_m36_default(); } }
void set_referencing_m36(ExcelValue newValue) { variable_set[106] = 1; referencing_m36_variable = newValue; }

ExcelValue referencing_n36_default() {
  return constant64;
}
static ExcelValue referencing_n36_variable;
ExcelValue referencing_n36() { if(variable_set[107] == 1) { return referencing_n36_variable; } else { return referencing_n36_default(); } }
void set_referencing_n36(ExcelValue newValue) { variable_set[107] = 1; referencing_n36_variable = newValue; }

ExcelValue referencing_o36() {
  static ExcelValue result;
  if(variable_set[108] == 1) { return result;}
  result = constant61;
  variable_set[108] = 1;
  return result;
}

ExcelValue referencing_c37() {
  static ExcelValue result;
  if(variable_set[109] == 1) { return result;}
  result = constant18;
  variable_set[109] = 1;
  return result;
}

ExcelValue referencing_d37() {
  static ExcelValue result;
  if(variable_set[110] == 1) { return result;}
  result = constant65;
  variable_set[110] = 1;
  return result;
}

ExcelValue referencing_f37_default() {
  return constant58;
}
static ExcelValue referencing_f37_variable;
ExcelValue referencing_f37() { if(variable_set[111] == 1) { return referencing_f37_variable; } else { return referencing_f37_default(); } }
void set_referencing_f37(ExcelValue newValue) { variable_set[111] = 1; referencing_f37_variable = newValue; }

ExcelValue referencing_m37_default() {
  return constant2;
}
static ExcelValue referencing_m37_variable;
ExcelValue referencing_m37() { if(variable_set[112] == 1) { return referencing_m37_variable; } else { return referencing_m37_default(); } }
void set_referencing_m37(ExcelValue newValue) { variable_set[112] = 1; referencing_m37_variable = newValue; }

ExcelValue referencing_n37_default() {
  return constant61;
}
static ExcelValue referencing_n37_variable;
ExcelValue referencing_n37() { if(variable_set[113] == 1) { return referencing_n37_variable; } else { return referencing_n37_default(); } }
void set_referencing_n37(ExcelValue newValue) { variable_set[113] = 1; referencing_n37_variable = newValue; }

ExcelValue referencing_o37() {
  static ExcelValue result;
  if(variable_set[114] == 1) { return result;}
  result = constant61;
  variable_set[114] = 1;
  return result;
}

ExcelValue referencing_c38() {
  static ExcelValue result;
  if(variable_set[115] == 1) { return result;}
  result = constant22;
  variable_set[115] = 1;
  return result;
}

ExcelValue referencing_d38() {
  static ExcelValue result;
  if(variable_set[116] == 1) { return result;}
  result = constant66;
  variable_set[116] = 1;
  return result;
}

ExcelValue referencing_i38_default() {
  return constant58;
}
static ExcelValue referencing_i38_variable;
ExcelValue referencing_i38() { if(variable_set[117] == 1) { return referencing_i38_variable; } else { return referencing_i38_default(); } }
void set_referencing_i38(ExcelValue newValue) { variable_set[117] = 1; referencing_i38_variable = newValue; }

ExcelValue referencing_m38_default() {
  return constant67;
}
static ExcelValue referencing_m38_variable;
ExcelValue referencing_m38() { if(variable_set[118] == 1) { return referencing_m38_variable; } else { return referencing_m38_default(); } }
void set_referencing_m38(ExcelValue newValue) { variable_set[118] = 1; referencing_m38_variable = newValue; }

ExcelValue referencing_n38_default() {
  return constant68;
}
static ExcelValue referencing_n38_variable;
ExcelValue referencing_n38() { if(variable_set[119] == 1) { return referencing_n38_variable; } else { return referencing_n38_default(); } }
void set_referencing_n38(ExcelValue newValue) { variable_set[119] = 1; referencing_n38_variable = newValue; }

ExcelValue referencing_o38() {
  static ExcelValue result;
  if(variable_set[120] == 1) { return result;}
  result = constant69;
  variable_set[120] = 1;
  return result;
}

ExcelValue referencing_c39() {
  static ExcelValue result;
  if(variable_set[121] == 1) { return result;}
  result = constant70;
  variable_set[121] = 1;
  return result;
}

ExcelValue referencing_d39() {
  static ExcelValue result;
  if(variable_set[122] == 1) { return result;}
  result = constant71;
  variable_set[122] = 1;
  return result;
}

ExcelValue referencing_e39() {
  static ExcelValue result;
  if(variable_set[123] == 1) { return result;}
  result = constant72;
  variable_set[123] = 1;
  return result;
}

ExcelValue referencing_h39_default() {
  return constant58;
}
static ExcelValue referencing_h39_variable;
ExcelValue referencing_h39() { if(variable_set[124] == 1) { return referencing_h39_variable; } else { return referencing_h39_default(); } }
void set_referencing_h39(ExcelValue newValue) { variable_set[124] = 1; referencing_h39_variable = newValue; }

ExcelValue referencing_m39_default() {
  return constant73;
}
static ExcelValue referencing_m39_variable;
ExcelValue referencing_m39() { if(variable_set[125] == 1) { return referencing_m39_variable; } else { return referencing_m39_default(); } }
void set_referencing_m39(ExcelValue newValue) { variable_set[125] = 1; referencing_m39_variable = newValue; }

ExcelValue referencing_n39_default() {
  return constant74;
}
static ExcelValue referencing_n39_variable;
ExcelValue referencing_n39() { if(variable_set[126] == 1) { return referencing_n39_variable; } else { return referencing_n39_default(); } }
void set_referencing_n39(ExcelValue newValue) { variable_set[126] = 1; referencing_n39_variable = newValue; }

ExcelValue referencing_o39() {
  static ExcelValue result;
  if(variable_set[127] == 1) { return result;}
  result = constant61;
  variable_set[127] = 1;
  return result;
}

ExcelValue referencing_c40() {
  static ExcelValue result;
  if(variable_set[128] == 1) { return result;}
  result = constant75;
  variable_set[128] = 1;
  return result;
}

ExcelValue referencing_d40() {
  static ExcelValue result;
  if(variable_set[129] == 1) { return result;}
  result = constant76;
  variable_set[129] = 1;
  return result;
}

ExcelValue referencing_e40() {
  static ExcelValue result;
  if(variable_set[130] == 1) { return result;}
  result = constant77;
  variable_set[130] = 1;
  return result;
}

ExcelValue referencing_g40_default() {
  return constant78;
}
static ExcelValue referencing_g40_variable;
ExcelValue referencing_g40() { if(variable_set[131] == 1) { return referencing_g40_variable; } else { return referencing_g40_default(); } }
void set_referencing_g40(ExcelValue newValue) { variable_set[131] = 1; referencing_g40_variable = newValue; }

ExcelValue referencing_j40_default() {
  return constant58;
}
static ExcelValue referencing_j40_variable;
ExcelValue referencing_j40() { if(variable_set[132] == 1) { return referencing_j40_variable; } else { return referencing_j40_default(); } }
void set_referencing_j40(ExcelValue newValue) { variable_set[132] = 1; referencing_j40_variable = newValue; }

ExcelValue referencing_m40_default() {
  return constant79;
}
static ExcelValue referencing_m40_variable;
ExcelValue referencing_m40() { if(variable_set[133] == 1) { return referencing_m40_variable; } else { return referencing_m40_default(); } }
void set_referencing_m40(ExcelValue newValue) { variable_set[133] = 1; referencing_m40_variable = newValue; }

ExcelValue referencing_n40_default() {
  return constant80;
}
static ExcelValue referencing_n40_variable;
ExcelValue referencing_n40() { if(variable_set[134] == 1) { return referencing_n40_variable; } else { return referencing_n40_default(); } }
void set_referencing_n40(ExcelValue newValue) { variable_set[134] = 1; referencing_n40_variable = newValue; }

ExcelValue referencing_o40() {
  static ExcelValue result;
  if(variable_set[135] == 1) { return result;}
  result = constant61;
  variable_set[135] = 1;
  return result;
}

ExcelValue referencing_c41() {
  static ExcelValue result;
  if(variable_set[136] == 1) { return result;}
  result = constant81;
  variable_set[136] = 1;
  return result;
}

ExcelValue referencing_d41() {
  static ExcelValue result;
  if(variable_set[137] == 1) { return result;}
  result = constant82;
  variable_set[137] = 1;
  return result;
}

ExcelValue referencing_e41() {
  static ExcelValue result;
  if(variable_set[138] == 1) { return result;}
  result = constant77;
  variable_set[138] = 1;
  return result;
}

ExcelValue referencing_g41_default() {
  return constant83;
}
static ExcelValue referencing_g41_variable;
ExcelValue referencing_g41() { if(variable_set[139] == 1) { return referencing_g41_variable; } else { return referencing_g41_default(); } }
void set_referencing_g41(ExcelValue newValue) { variable_set[139] = 1; referencing_g41_variable = newValue; }

ExcelValue referencing_j41_default() {
  return constant58;
}
static ExcelValue referencing_j41_variable;
ExcelValue referencing_j41() { if(variable_set[140] == 1) { return referencing_j41_variable; } else { return referencing_j41_default(); } }
void set_referencing_j41(ExcelValue newValue) { variable_set[140] = 1; referencing_j41_variable = newValue; }

ExcelValue referencing_m41_default() {
  return constant83;
}
static ExcelValue referencing_m41_variable;
ExcelValue referencing_m41() { if(variable_set[141] == 1) { return referencing_m41_variable; } else { return referencing_m41_default(); } }
void set_referencing_m41(ExcelValue newValue) { variable_set[141] = 1; referencing_m41_variable = newValue; }

ExcelValue referencing_n41_default() {
  return constant84;
}
static ExcelValue referencing_n41_variable;
ExcelValue referencing_n41() { if(variable_set[142] == 1) { return referencing_n41_variable; } else { return referencing_n41_default(); } }
void set_referencing_n41(ExcelValue newValue) { variable_set[142] = 1; referencing_n41_variable = newValue; }

ExcelValue referencing_o41() {
  static ExcelValue result;
  if(variable_set[143] == 1) { return result;}
  result = constant61;
  variable_set[143] = 1;
  return result;
}

ExcelValue referencing_c42() {
  static ExcelValue result;
  if(variable_set[144] == 1) { return result;}
  result = constant85;
  variable_set[144] = 1;
  return result;
}

ExcelValue referencing_d42() {
  static ExcelValue result;
  if(variable_set[145] == 1) { return result;}
  result = constant86;
  variable_set[145] = 1;
  return result;
}

ExcelValue referencing_f42_default() {
  return constant58;
}
static ExcelValue referencing_f42_variable;
ExcelValue referencing_f42() { if(variable_set[146] == 1) { return referencing_f42_variable; } else { return referencing_f42_default(); } }
void set_referencing_f42(ExcelValue newValue) { variable_set[146] = 1; referencing_f42_variable = newValue; }

ExcelValue referencing_l42_default() {
  return constant58;
}
static ExcelValue referencing_l42_variable;
ExcelValue referencing_l42() { if(variable_set[147] == 1) { return referencing_l42_variable; } else { return referencing_l42_default(); } }
void set_referencing_l42(ExcelValue newValue) { variable_set[147] = 1; referencing_l42_variable = newValue; }

ExcelValue referencing_m42_default() {
  return constant5;
}
static ExcelValue referencing_m42_variable;
ExcelValue referencing_m42() { if(variable_set[148] == 1) { return referencing_m42_variable; } else { return referencing_m42_default(); } }
void set_referencing_m42(ExcelValue newValue) { variable_set[148] = 1; referencing_m42_variable = newValue; }

ExcelValue referencing_o42() {
  static ExcelValue result;
  if(variable_set[149] == 1) { return result;}
  result = constant61;
  variable_set[149] = 1;
  return result;
}

ExcelValue referencing_c43() {
  static ExcelValue result;
  if(variable_set[150] == 1) { return result;}
  result = constant87;
  variable_set[150] = 1;
  return result;
}

ExcelValue referencing_d43() {
  static ExcelValue result;
  if(variable_set[151] == 1) { return result;}
  result = constant88;
  variable_set[151] = 1;
  return result;
}

ExcelValue referencing_f43_default() {
  return constant58;
}
static ExcelValue referencing_f43_variable;
ExcelValue referencing_f43() { if(variable_set[152] == 1) { return referencing_f43_variable; } else { return referencing_f43_default(); } }
void set_referencing_f43(ExcelValue newValue) { variable_set[152] = 1; referencing_f43_variable = newValue; }

ExcelValue referencing_l43_default() {
  return constant89;
}
static ExcelValue referencing_l43_variable;
ExcelValue referencing_l43() { if(variable_set[153] == 1) { return referencing_l43_variable; } else { return referencing_l43_default(); } }
void set_referencing_l43(ExcelValue newValue) { variable_set[153] = 1; referencing_l43_variable = newValue; }

ExcelValue referencing_m43_default() {
  return constant18;
}
static ExcelValue referencing_m43_variable;
ExcelValue referencing_m43() { if(variable_set[154] == 1) { return referencing_m43_variable; } else { return referencing_m43_default(); } }
void set_referencing_m43(ExcelValue newValue) { variable_set[154] = 1; referencing_m43_variable = newValue; }

ExcelValue referencing_o43() {
  static ExcelValue result;
  if(variable_set[155] == 1) { return result;}
  result = constant61;
  variable_set[155] = 1;
  return result;
}

ExcelValue referencing_c44() {
  static ExcelValue result;
  if(variable_set[156] == 1) { return result;}
  result = constant19;
  variable_set[156] = 1;
  return result;
}

ExcelValue referencing_d44() {
  static ExcelValue result;
  if(variable_set[157] == 1) { return result;}
  result = constant90;
  variable_set[157] = 1;
  return result;
}

ExcelValue referencing_l44_default() {
  return constant58;
}
static ExcelValue referencing_l44_variable;
ExcelValue referencing_l44() { if(variable_set[158] == 1) { return referencing_l44_variable; } else { return referencing_l44_default(); } }
void set_referencing_l44(ExcelValue newValue) { variable_set[158] = 1; referencing_l44_variable = newValue; }

ExcelValue referencing_m44_default() {
  return constant91;
}
static ExcelValue referencing_m44_variable;
ExcelValue referencing_m44() { if(variable_set[159] == 1) { return referencing_m44_variable; } else { return referencing_m44_default(); } }
void set_referencing_m44(ExcelValue newValue) { variable_set[159] = 1; referencing_m44_variable = newValue; }

ExcelValue referencing_n44_default() {
  return constant92;
}
static ExcelValue referencing_n44_variable;
ExcelValue referencing_n44() { if(variable_set[160] == 1) { return referencing_n44_variable; } else { return referencing_n44_default(); } }
void set_referencing_n44(ExcelValue newValue) { variable_set[160] = 1; referencing_n44_variable = newValue; }

ExcelValue referencing_o44() {
  static ExcelValue result;
  if(variable_set[161] == 1) { return result;}
  result = constant61;
  variable_set[161] = 1;
  return result;
}

ExcelValue referencing_c45() {
  static ExcelValue result;
  if(variable_set[162] == 1) { return result;}
  result = constant93;
  variable_set[162] = 1;
  return result;
}

ExcelValue referencing_d45() {
  static ExcelValue result;
  if(variable_set[163] == 1) { return result;}
  result = constant94;
  variable_set[163] = 1;
  return result;
}

ExcelValue referencing_g45_default() {
  return constant95;
}
static ExcelValue referencing_g45_variable;
ExcelValue referencing_g45() { if(variable_set[164] == 1) { return referencing_g45_variable; } else { return referencing_g45_default(); } }
void set_referencing_g45(ExcelValue newValue) { variable_set[164] = 1; referencing_g45_variable = newValue; }

ExcelValue referencing_j45_default() {
  return constant58;
}
static ExcelValue referencing_j45_variable;
ExcelValue referencing_j45() { if(variable_set[165] == 1) { return referencing_j45_variable; } else { return referencing_j45_default(); } }
void set_referencing_j45(ExcelValue newValue) { variable_set[165] = 1; referencing_j45_variable = newValue; }

ExcelValue referencing_m45_default() {
  return constant95;
}
static ExcelValue referencing_m45_variable;
ExcelValue referencing_m45() { if(variable_set[166] == 1) { return referencing_m45_variable; } else { return referencing_m45_default(); } }
void set_referencing_m45(ExcelValue newValue) { variable_set[166] = 1; referencing_m45_variable = newValue; }

ExcelValue referencing_n45_default() {
  return constant60;
}
static ExcelValue referencing_n45_variable;
ExcelValue referencing_n45() { if(variable_set[167] == 1) { return referencing_n45_variable; } else { return referencing_n45_default(); } }
void set_referencing_n45(ExcelValue newValue) { variable_set[167] = 1; referencing_n45_variable = newValue; }

ExcelValue referencing_o45() {
  static ExcelValue result;
  if(variable_set[168] == 1) { return result;}
  result = constant61;
  variable_set[168] = 1;
  return result;
}

ExcelValue referencing_c46() {
  static ExcelValue result;
  if(variable_set[169] == 1) { return result;}
  result = constant27;
  variable_set[169] = 1;
  return result;
}

ExcelValue referencing_d46() {
  static ExcelValue result;
  if(variable_set[170] == 1) { return result;}
  result = constant96;
  variable_set[170] = 1;
  return result;
}

ExcelValue referencing_g46_default() {
  return constant97;
}
static ExcelValue referencing_g46_variable;
ExcelValue referencing_g46() { if(variable_set[171] == 1) { return referencing_g46_variable; } else { return referencing_g46_default(); } }
void set_referencing_g46(ExcelValue newValue) { variable_set[171] = 1; referencing_g46_variable = newValue; }

ExcelValue referencing_h46_default() {
  return constant58;
}
static ExcelValue referencing_h46_variable;
ExcelValue referencing_h46() { if(variable_set[172] == 1) { return referencing_h46_variable; } else { return referencing_h46_default(); } }
void set_referencing_h46(ExcelValue newValue) { variable_set[172] = 1; referencing_h46_variable = newValue; }

ExcelValue referencing_m46_default() {
  return constant98;
}
static ExcelValue referencing_m46_variable;
ExcelValue referencing_m46() { if(variable_set[173] == 1) { return referencing_m46_variable; } else { return referencing_m46_default(); } }
void set_referencing_m46(ExcelValue newValue) { variable_set[173] = 1; referencing_m46_variable = newValue; }

ExcelValue referencing_n46_default() {
  return constant99;
}
static ExcelValue referencing_n46_variable;
ExcelValue referencing_n46() { if(variable_set[174] == 1) { return referencing_n46_variable; } else { return referencing_n46_default(); } }
void set_referencing_n46(ExcelValue newValue) { variable_set[174] = 1; referencing_n46_variable = newValue; }

ExcelValue referencing_o46() {
  static ExcelValue result;
  if(variable_set[175] == 1) { return result;}
  result = constant61;
  variable_set[175] = 1;
  return result;
}

ExcelValue referencing_c47() {
  static ExcelValue result;
  if(variable_set[176] == 1) { return result;}
  result = constant100;
  variable_set[176] = 1;
  return result;
}

ExcelValue referencing_d47() {
  static ExcelValue result;
  if(variable_set[177] == 1) { return result;}
  result = constant101;
  variable_set[177] = 1;
  return result;
}

ExcelValue referencing_e47() {
  static ExcelValue result;
  if(variable_set[178] == 1) { return result;}
  result = constant102;
  variable_set[178] = 1;
  return result;
}

ExcelValue referencing_k47_default() {
  return constant58;
}
static ExcelValue referencing_k47_variable;
ExcelValue referencing_k47() { if(variable_set[179] == 1) { return referencing_k47_variable; } else { return referencing_k47_default(); } }
void set_referencing_k47(ExcelValue newValue) { variable_set[179] = 1; referencing_k47_variable = newValue; }

ExcelValue referencing_m47_default() {
  return constant103;
}
static ExcelValue referencing_m47_variable;
ExcelValue referencing_m47() { if(variable_set[180] == 1) { return referencing_m47_variable; } else { return referencing_m47_default(); } }
void set_referencing_m47(ExcelValue newValue) { variable_set[180] = 1; referencing_m47_variable = newValue; }

ExcelValue referencing_n47_default() {
  return constant84;
}
static ExcelValue referencing_n47_variable;
ExcelValue referencing_n47() { if(variable_set[181] == 1) { return referencing_n47_variable; } else { return referencing_n47_default(); } }
void set_referencing_n47(ExcelValue newValue) { variable_set[181] = 1; referencing_n47_variable = newValue; }

ExcelValue referencing_o47() {
  static ExcelValue result;
  if(variable_set[182] == 1) { return result;}
  result = constant61;
  variable_set[182] = 1;
  return result;
}

ExcelValue referencing_d50() {
  static ExcelValue result;
  if(variable_set[183] == 1) { return result;}
  result = constant57;
  variable_set[183] = 1;
  return result;
}

ExcelValue referencing_g50_default() {
  return constant104;
}
static ExcelValue referencing_g50_variable;
ExcelValue referencing_g50() { if(variable_set[184] == 1) { return referencing_g50_variable; } else { return referencing_g50_default(); } }
void set_referencing_g50(ExcelValue newValue) { variable_set[184] = 1; referencing_g50_variable = newValue; }

ExcelValue referencing_d51() {
  static ExcelValue result;
  if(variable_set[185] == 1) { return result;}
  result = constant62;
  variable_set[185] = 1;
  return result;
}

ExcelValue referencing_g51_default() {
  return constant105;
}
static ExcelValue referencing_g51_variable;
ExcelValue referencing_g51() { if(variable_set[186] == 1) { return referencing_g51_variable; } else { return referencing_g51_default(); } }
void set_referencing_g51(ExcelValue newValue) { variable_set[186] = 1; referencing_g51_variable = newValue; }

ExcelValue referencing_d52() {
  static ExcelValue result;
  if(variable_set[187] == 1) { return result;}
  result = constant65;
  variable_set[187] = 1;
  return result;
}

ExcelValue referencing_g52_default() {
  return constant106;
}
static ExcelValue referencing_g52_variable;
ExcelValue referencing_g52() { if(variable_set[188] == 1) { return referencing_g52_variable; } else { return referencing_g52_default(); } }
void set_referencing_g52(ExcelValue newValue) { variable_set[188] = 1; referencing_g52_variable = newValue; }

ExcelValue referencing_d53() {
  static ExcelValue result;
  if(variable_set[189] == 1) { return result;}
  result = constant66;
  variable_set[189] = 1;
  return result;
}

ExcelValue referencing_g53_default() {
  return constant107;
}
static ExcelValue referencing_g53_variable;
ExcelValue referencing_g53() { if(variable_set[190] == 1) { return referencing_g53_variable; } else { return referencing_g53_default(); } }
void set_referencing_g53(ExcelValue newValue) { variable_set[190] = 1; referencing_g53_variable = newValue; }

ExcelValue referencing_d54() {
  static ExcelValue result;
  if(variable_set[191] == 1) { return result;}
  result = constant71;
  variable_set[191] = 1;
  return result;
}

ExcelValue referencing_g54_default() {
  return constant107;
}
static ExcelValue referencing_g54_variable;
ExcelValue referencing_g54() { if(variable_set[192] == 1) { return referencing_g54_variable; } else { return referencing_g54_default(); } }
void set_referencing_g54(ExcelValue newValue) { variable_set[192] = 1; referencing_g54_variable = newValue; }

ExcelValue referencing_d55() {
  static ExcelValue result;
  if(variable_set[193] == 1) { return result;}
  result = constant76;
  variable_set[193] = 1;
  return result;
}

ExcelValue referencing_g55_default() {
  return constant61;
}
static ExcelValue referencing_g55_variable;
ExcelValue referencing_g55() { if(variable_set[194] == 1) { return referencing_g55_variable; } else { return referencing_g55_default(); } }
void set_referencing_g55(ExcelValue newValue) { variable_set[194] = 1; referencing_g55_variable = newValue; }

ExcelValue referencing_d56() {
  static ExcelValue result;
  if(variable_set[195] == 1) { return result;}
  result = constant82;
  variable_set[195] = 1;
  return result;
}

ExcelValue referencing_g56_default() {
  return constant61;
}
static ExcelValue referencing_g56_variable;
ExcelValue referencing_g56() { if(variable_set[196] == 1) { return referencing_g56_variable; } else { return referencing_g56_default(); } }
void set_referencing_g56(ExcelValue newValue) { variable_set[196] = 1; referencing_g56_variable = newValue; }

ExcelValue referencing_d57() {
  static ExcelValue result;
  if(variable_set[197] == 1) { return result;}
  result = constant86;
  variable_set[197] = 1;
  return result;
}

ExcelValue referencing_g57_default() {
  return constant61;
}
static ExcelValue referencing_g57_variable;
ExcelValue referencing_g57() { if(variable_set[198] == 1) { return referencing_g57_variable; } else { return referencing_g57_default(); } }
void set_referencing_g57(ExcelValue newValue) { variable_set[198] = 1; referencing_g57_variable = newValue; }

ExcelValue referencing_d58() {
  static ExcelValue result;
  if(variable_set[199] == 1) { return result;}
  result = constant88;
  variable_set[199] = 1;
  return result;
}

ExcelValue referencing_g58_default() {
  return constant61;
}
static ExcelValue referencing_g58_variable;
ExcelValue referencing_g58() { if(variable_set[200] == 1) { return referencing_g58_variable; } else { return referencing_g58_default(); } }
void set_referencing_g58(ExcelValue newValue) { variable_set[200] = 1; referencing_g58_variable = newValue; }

ExcelValue referencing_d59() {
  static ExcelValue result;
  if(variable_set[201] == 1) { return result;}
  result = constant90;
  variable_set[201] = 1;
  return result;
}

ExcelValue referencing_g59_default() {
  return constant61;
}
static ExcelValue referencing_g59_variable;
ExcelValue referencing_g59() { if(variable_set[202] == 1) { return referencing_g59_variable; } else { return referencing_g59_default(); } }
void set_referencing_g59(ExcelValue newValue) { variable_set[202] = 1; referencing_g59_variable = newValue; }

ExcelValue referencing_d60() {
  static ExcelValue result;
  if(variable_set[203] == 1) { return result;}
  result = constant94;
  variable_set[203] = 1;
  return result;
}

ExcelValue referencing_g60_default() {
  return constant61;
}
static ExcelValue referencing_g60_variable;
ExcelValue referencing_g60() { if(variable_set[204] == 1) { return referencing_g60_variable; } else { return referencing_g60_default(); } }
void set_referencing_g60(ExcelValue newValue) { variable_set[204] = 1; referencing_g60_variable = newValue; }

ExcelValue referencing_d61() {
  static ExcelValue result;
  if(variable_set[205] == 1) { return result;}
  result = constant96;
  variable_set[205] = 1;
  return result;
}

ExcelValue referencing_g61_default() {
  return constant61;
}
static ExcelValue referencing_g61_variable;
ExcelValue referencing_g61() { if(variable_set[206] == 1) { return referencing_g61_variable; } else { return referencing_g61_default(); } }
void set_referencing_g61(ExcelValue newValue) { variable_set[206] = 1; referencing_g61_variable = newValue; }

ExcelValue referencing_d62() {
  static ExcelValue result;
  if(variable_set[207] == 1) { return result;}
  result = constant101;
  variable_set[207] = 1;
  return result;
}

ExcelValue referencing_g62_default() {
  return constant61;
}
static ExcelValue referencing_g62_variable;
ExcelValue referencing_g62() { if(variable_set[208] == 1) { return referencing_g62_variable; } else { return referencing_g62_default(); } }
void set_referencing_g62(ExcelValue newValue) { variable_set[208] = 1; referencing_g62_variable = newValue; }

ExcelValue referencing_d64_default() {
  return constant55;
}
static ExcelValue referencing_d64_variable;
ExcelValue referencing_d64() { if(variable_set[209] == 1) { return referencing_d64_variable; } else { return referencing_d64_default(); } }
void set_referencing_d64(ExcelValue newValue) { variable_set[209] = 1; referencing_d64_variable = newValue; }

ExcelValue referencing_e64() {
  static ExcelValue result;
  if(variable_set[210] == 1) { return result;}
  result = constant42;
  variable_set[210] = 1;
  return result;
}

ExcelValue referencing_h64() {
  static ExcelValue result;
  if(variable_set[211] == 1) { return result;}
  static ExcelValue array1[117];
  array1[0] = constant108;
  array1[1] = constant108;
  array1[2] = constant108;
  array1[3] = constant108;
  array1[4] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_j34())),referencing_j35());
  array1[5] = constant108;
  array1[6] = constant108;
  array1[7] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_m34())),referencing_m35());
  array1[8] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_n34())),referencing_n35());
  array1[9] = constant108;
  array1[10] = constant108;
  array1[11] = constant108;
  array1[12] = constant108;
  array1[13] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_j34())),referencing_j36());
  array1[14] = constant108;
  array1[15] = constant108;
  array1[16] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_m34())),referencing_m36());
  array1[17] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_n34())),referencing_n36());
  array1[18] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_f34())),referencing_f37());
  array1[19] = constant108;
  array1[20] = constant108;
  array1[21] = constant108;
  array1[22] = constant108;
  array1[23] = constant108;
  array1[24] = constant108;
  array1[25] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_m34())),referencing_m37());
  array1[26] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_n34())),referencing_n37());
  array1[27] = constant108;
  array1[28] = constant108;
  array1[29] = constant108;
  array1[30] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_i34())),referencing_i38());
  array1[31] = constant108;
  array1[32] = constant108;
  array1[33] = constant108;
  array1[34] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_m34())),referencing_m38());
  array1[35] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_n34())),referencing_n38());
  array1[36] = constant108;
  array1[37] = constant108;
  array1[38] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_h34())),referencing_h39());
  array1[39] = constant108;
  array1[40] = constant108;
  array1[41] = constant108;
  array1[42] = constant108;
  array1[43] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_m34())),referencing_m39());
  array1[44] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_n34())),referencing_n39());
  array1[45] = constant108;
  array1[46] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_g34())),referencing_g40());
  array1[47] = constant108;
  array1[48] = constant108;
  array1[49] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_j34())),referencing_j40());
  array1[50] = constant108;
  array1[51] = constant108;
  array1[52] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_m34())),referencing_m40());
  array1[53] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_n34())),referencing_n40());
  array1[54] = constant108;
  array1[55] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_g34())),referencing_g41());
  array1[56] = constant108;
  array1[57] = constant108;
  array1[58] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_j34())),referencing_j41());
  array1[59] = constant108;
  array1[60] = constant108;
  array1[61] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_m34())),referencing_m41());
  array1[62] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_n34())),referencing_n41());
  array1[63] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_f34())),referencing_f42());
  array1[64] = constant108;
  array1[65] = constant108;
  array1[66] = constant108;
  array1[67] = constant108;
  array1[68] = constant108;
  array1[69] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_l34())),referencing_l42());
  array1[70] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_m34())),referencing_m42());
  array1[71] = constant108;
  array1[72] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_f34())),referencing_f43());
  array1[73] = constant108;
  array1[74] = constant108;
  array1[75] = constant108;
  array1[76] = constant108;
  array1[77] = constant108;
  array1[78] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_l34())),referencing_l43());
  array1[79] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_m34())),referencing_m43());
  array1[80] = constant108;
  array1[81] = constant108;
  array1[82] = constant108;
  array1[83] = constant108;
  array1[84] = constant108;
  array1[85] = constant108;
  array1[86] = constant108;
  array1[87] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_l34())),referencing_l44());
  array1[88] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_m34())),referencing_m44());
  array1[89] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_n34())),referencing_n44());
  array1[90] = constant108;
  array1[91] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_g34())),referencing_g45());
  array1[92] = constant108;
  array1[93] = constant108;
  array1[94] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_j34())),referencing_j45());
  array1[95] = constant108;
  array1[96] = constant108;
  array1[97] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_m34())),referencing_m45());
  array1[98] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_n34())),referencing_n45());
  array1[99] = constant108;
  array1[100] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_g34())),referencing_g46());
  array1[101] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_h34())),referencing_h46());
  array1[102] = constant108;
  array1[103] = constant108;
  array1[104] = constant108;
  array1[105] = constant108;
  array1[106] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_m34())),referencing_m46());
  array1[107] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_n34())),referencing_n46());
  array1[108] = constant108;
  array1[109] = constant108;
  array1[110] = constant108;
  array1[111] = constant108;
  array1[112] = constant108;
  array1[113] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_k34())),referencing_k47());
  array1[114] = constant108;
  array1[115] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_m34())),referencing_m47());
  array1[116] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_n34())),referencing_n47());
  ExcelValue array1_ev = new_excel_range(array1,13,9);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[211] = 1;
  return result;
}

ExcelValue tables_a1() {
  static ExcelValue result;
  if(variable_set[212] == 1) { return result;}
  result = BLANK;
  variable_set[212] = 1;
  return result;
}

ExcelValue tables_b2_default() {
  return constant109;
}
static ExcelValue tables_b2_variable;
ExcelValue tables_b2() { if(variable_set[213] == 1) { return tables_b2_variable; } else { return tables_b2_default(); } }
void set_tables_b2(ExcelValue newValue) { variable_set[213] = 1; tables_b2_variable = newValue; }

ExcelValue tables_c2_default() {
  return constant110;
}
static ExcelValue tables_c2_variable;
ExcelValue tables_c2() { if(variable_set[214] == 1) { return tables_c2_variable; } else { return tables_c2_default(); } }
void set_tables_c2(ExcelValue newValue) { variable_set[214] = 1; tables_c2_variable = newValue; }

ExcelValue tables_d2_default() {
  return constant111;
}
static ExcelValue tables_d2_variable;
ExcelValue tables_d2() { if(variable_set[215] == 1) { return tables_d2_variable; } else { return tables_d2_default(); } }
void set_tables_d2(ExcelValue newValue) { variable_set[215] = 1; tables_d2_variable = newValue; }

ExcelValue tables_b3_default() {
  return constant2;
}
static ExcelValue tables_b3_variable;
ExcelValue tables_b3() { if(variable_set[216] == 1) { return tables_b3_variable; } else { return tables_b3_default(); } }
void set_tables_b3(ExcelValue newValue) { variable_set[216] = 1; tables_b3_variable = newValue; }

ExcelValue tables_c3_default() {
  return constant112;
}
static ExcelValue tables_c3_variable;
ExcelValue tables_c3() { if(variable_set[217] == 1) { return tables_c3_variable; } else { return tables_c3_default(); } }
void set_tables_c3(ExcelValue newValue) { variable_set[217] = 1; tables_c3_variable = newValue; }

ExcelValue tables_d3() {
  static ExcelValue result;
  if(variable_set[218] == 1) { return result;}
  ExcelValue array0[] = {tables_b3(),tables_c3()};
  result = string_join(2, array0);
  variable_set[218] = 1;
  return result;
}

ExcelValue tables_b4_default() {
  return constant5;
}
static ExcelValue tables_b4_variable;
ExcelValue tables_b4() { if(variable_set[219] == 1) { return tables_b4_variable; } else { return tables_b4_default(); } }
void set_tables_b4(ExcelValue newValue) { variable_set[219] = 1; tables_b4_variable = newValue; }

ExcelValue tables_c4_default() {
  return constant113;
}
static ExcelValue tables_c4_variable;
ExcelValue tables_c4() { if(variable_set[220] == 1) { return tables_c4_variable; } else { return tables_c4_default(); } }
void set_tables_c4(ExcelValue newValue) { variable_set[220] = 1; tables_c4_variable = newValue; }

ExcelValue tables_d4() {
  static ExcelValue result;
  if(variable_set[221] == 1) { return result;}
  ExcelValue array0[] = {tables_b4(),tables_c4()};
  result = string_join(2, array0);
  variable_set[221] = 1;
  return result;
}

ExcelValue tables_f4() {
  static ExcelValue result;
  if(variable_set[222] == 1) { return result;}
  result = tables_c4();
  variable_set[222] = 1;
  return result;
}

ExcelValue tables_g4() {
  static ExcelValue result;
  if(variable_set[223] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = tables_b4();
  array0[1] = tables_c4();
  array0[2] = tables_d4();
  ExcelValue array0_ev = new_excel_range(array0,1,3);
  result = excel_match(constant114,array0_ev,FALSE);
  variable_set[223] = 1;
  return result;
}

ExcelValue tables_h4() {
  static ExcelValue result;
  if(variable_set[224] == 1) { return result;}
  static ExcelValue array0[2];
  array0[0] = tables_c4();
  array0[1] = tables_d4();
  ExcelValue array0_ev = new_excel_range(array0,1,2);
  result = excel_match_2(constant113,array0_ev);
  variable_set[224] = 1;
  return result;
}

ExcelValue tables_b5() {
  static ExcelValue result;
  if(variable_set[225] == 1) { return result;}
  result = _common7();
  variable_set[225] = 1;
  return result;
}

ExcelValue tables_c5() {
  static ExcelValue result;
  if(variable_set[226] == 1) { return result;}
  static ExcelValue array1[2];
  array1[0] = tables_c3();
  array1[1] = tables_c4();
  ExcelValue array1_ev = new_excel_range(array1,2,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[226] = 1;
  return result;
}

ExcelValue tables_e6() {
  static ExcelValue result;
  if(variable_set[227] == 1) { return result;}
  result = tables_b2();
  variable_set[227] = 1;
  return result;
}

ExcelValue tables_f6() {
  static ExcelValue result;
  if(variable_set[228] == 1) { return result;}
  result = tables_c2();
  variable_set[228] = 1;
  return result;
}

ExcelValue tables_g6() {
  static ExcelValue result;
  if(variable_set[229] == 1) { return result;}
  result = tables_d2();
  variable_set[229] = 1;
  return result;
}

ExcelValue tables_e7() {
  static ExcelValue result;
  if(variable_set[230] == 1) { return result;}
  result = tables_b5();
  variable_set[230] = 1;
  return result;
}

ExcelValue tables_f7() {
  static ExcelValue result;
  if(variable_set[231] == 1) { return result;}
  result = tables_c5();
  variable_set[231] = 1;
  return result;
}

ExcelValue tables_g7() {
  static ExcelValue result;
  if(variable_set[232] == 1) { return result;}
  result = BLANK;
  variable_set[232] = 1;
  return result;
}

ExcelValue tables_e8() {
  static ExcelValue result;
  if(variable_set[233] == 1) { return result;}
  result = tables_b2();
  variable_set[233] = 1;
  return result;
}

ExcelValue tables_f8() {
  static ExcelValue result;
  if(variable_set[234] == 1) { return result;}
  result = tables_c2();
  variable_set[234] = 1;
  return result;
}

ExcelValue tables_g8() {
  static ExcelValue result;
  if(variable_set[235] == 1) { return result;}
  result = tables_d2();
  variable_set[235] = 1;
  return result;
}

ExcelValue tables_e9() {
  static ExcelValue result;
  if(variable_set[236] == 1) { return result;}
  result = tables_b3();
  variable_set[236] = 1;
  return result;
}

ExcelValue tables_f9() {
  static ExcelValue result;
  if(variable_set[237] == 1) { return result;}
  result = tables_c3();
  variable_set[237] = 1;
  return result;
}

ExcelValue tables_g9() {
  static ExcelValue result;
  if(variable_set[238] == 1) { return result;}
  result = tables_d3();
  variable_set[238] = 1;
  return result;
}

ExcelValue tables_c10() {
  static ExcelValue result;
  if(variable_set[239] == 1) { return result;}
  result = _common1();
  variable_set[239] = 1;
  return result;
}

ExcelValue tables_e10() {
  static ExcelValue result;
  if(variable_set[240] == 1) { return result;}
  result = tables_b4();
  variable_set[240] = 1;
  return result;
}

ExcelValue tables_f10() {
  static ExcelValue result;
  if(variable_set[241] == 1) { return result;}
  result = tables_c4();
  variable_set[241] = 1;
  return result;
}

ExcelValue tables_g10() {
  static ExcelValue result;
  if(variable_set[242] == 1) { return result;}
  result = tables_d4();
  variable_set[242] = 1;
  return result;
}

ExcelValue tables_c11() {
  static ExcelValue result;
  if(variable_set[243] == 1) { return result;}
  result = _common7();
  variable_set[243] = 1;
  return result;
}

ExcelValue tables_e11() {
  static ExcelValue result;
  if(variable_set[244] == 1) { return result;}
  result = tables_b5();
  variable_set[244] = 1;
  return result;
}

ExcelValue tables_f11() {
  static ExcelValue result;
  if(variable_set[245] == 1) { return result;}
  result = tables_c5();
  variable_set[245] = 1;
  return result;
}

ExcelValue tables_g11() {
  static ExcelValue result;
  if(variable_set[246] == 1) { return result;}
  result = BLANK;
  variable_set[246] = 1;
  return result;
}

ExcelValue tables_c12() {
  static ExcelValue result;
  if(variable_set[247] == 1) { return result;}
  result = tables_b5();
  variable_set[247] = 1;
  return result;
}

ExcelValue tables_c13() {
  static ExcelValue result;
  if(variable_set[248] == 1) { return result;}
  result = _common9();
  variable_set[248] = 1;
  return result;
}

ExcelValue tables_c14() {
  static ExcelValue result;
  if(variable_set[249] == 1) { return result;}
  result = _common9();
  variable_set[249] = 1;
  return result;
}

ExcelValue s_innapropriate_sheet_name__c4() {
  static ExcelValue result;
  if(variable_set[250] == 1) { return result;}
  result = valuetypes_a3();
  variable_set[250] = 1;
  return result;
}

ExcelValue _common0() {
  static ExcelValue result;
  if(variable_set[251] == 1) { return result;}
  result = add(referencing_a4(),constant2);
  variable_set[251] = 1;
  return result;
}

ExcelValue _common1() {
  static ExcelValue result;
  if(variable_set[252] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = tables_b5();
  array1[1] = tables_c5();
  array1[2] = BLANK;
  ExcelValue array1_ev = new_excel_range(array1,1,3);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[252] = 1;
  return result;
}

ExcelValue _common3() {
  static ExcelValue result;
  if(variable_set[253] == 1) { return result;}
  static ExcelValue array0[16];
  array0[0] = referencing_c16();
  array0[1] = referencing_d16();
  array0[2] = referencing_e16();
  array0[3] = referencing_f16();
  array0[4] = referencing_c17();
  array0[5] = referencing_d17();
  array0[6] = referencing_e17();
  array0[7] = referencing_f17();
  array0[8] = referencing_c18();
  array0[9] = referencing_d18();
  array0[10] = referencing_e18();
  array0[11] = referencing_f18();
  array0[12] = referencing_c19();
  array0[13] = referencing_d19();
  array0[14] = referencing_e19();
  array0[15] = referencing_f19();
  ExcelValue array0_ev = new_excel_range(array0,4,4);
  static ExcelValue array1[4];
  array1[0] = referencing_c15();
  array1[1] = referencing_d15();
  array1[2] = referencing_e15();
  array1[3] = referencing_f15();
  ExcelValue array1_ev = new_excel_range(array1,1,4);
  result = excel_index(array0_ev,BLANK,excel_match(referencing_c22(),array1_ev,constant61));
  variable_set[253] = 1;
  return result;
}

ExcelValue _common7() {
  static ExcelValue result;
  if(variable_set[254] == 1) { return result;}
  static ExcelValue array1[2];
  array1[0] = tables_b3();
  array1[1] = tables_b4();
  ExcelValue array1_ev = new_excel_range(array1,2,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[254] = 1;
  return result;
}

ExcelValue _common9() {
  static ExcelValue result;
  if(variable_set[255] == 1) { return result;}
  static ExcelValue array1[6];
  array1[0] = tables_b3();
  array1[1] = tables_c3();
  array1[2] = tables_d3();
  array1[3] = tables_b4();
  array1[4] = tables_c4();
  array1[5] = tables_d4();
  ExcelValue array1_ev = new_excel_range(array1,2,3);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[255] = 1;
  return result;
}

// Start of named references
// End of named references
