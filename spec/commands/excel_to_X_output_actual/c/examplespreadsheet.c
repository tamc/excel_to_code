// /Users/thomascounsell/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx approximately translated into C
// definitions
#define NUMBER_OF_REFS 349
#define EXCEL_FILENAME  "/Users/thomascounsell/Documents/github/excel_to_code/spec/test_data/ExampleSpreadsheet.xlsx"
// end of definitions

// First we have c versions of all the excel functions that we know
#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>
#include <math.h>
#include <locale.h>

#ifndef NUMBER_OF_REFS
  #define NUMBER_OF_REFS 0
#endif

#ifndef EXCEL_FILENAME
  #define EXCEL_FILENAME "NoExcelFilename"
#endif

// Need to retain malloc'd values for a while, so can return to functions that use this library
// So to avoid a memory leak we keep an array of all the values we have malloc'd, which we then
// free when the reset() function is called.
#define MEMORY_TO_BE_FREED_LATER_HEAP_INCREMENT 1000

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
static ExcelValue average(int array_size, ExcelValue *array);
static ExcelValue averageifs(ExcelValue average_range_v, int number_of_arguments, ExcelValue *arguments);
static ExcelValue excel_char(ExcelValue number_v);
static ExcelValue excel_ceiling_math_2(ExcelValue number_v, ExcelValue multiple_v);
static ExcelValue excel_ceiling_math(ExcelValue number_v, ExcelValue multiple_v, ExcelValue mode_v);
static ExcelValue ensure_is_number(ExcelValue maybe_number_v);
static ExcelValue find_2(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v);
static ExcelValue find(ExcelValue string_to_look_for_v, ExcelValue string_to_look_in_v, ExcelValue position_to_start_at_v);
static ExcelValue hlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v);
static ExcelValue hlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue row_number_v, ExcelValue match_type_v);
static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error);
static ExcelValue iserr(ExcelValue value);
static ExcelValue iserror(ExcelValue value);
static ExcelValue excel_index(ExcelValue array_v, ExcelValue row_number_v, ExcelValue column_number_v);
static ExcelValue excel_index_2(ExcelValue array_v, ExcelValue row_number_v);
static ExcelValue excel_isnumber(ExcelValue number);
static ExcelValue excel_isblank(ExcelValue value);
static ExcelValue forecast(ExcelValue required_x, ExcelValue known_y, ExcelValue known_x);
static ExcelValue large(ExcelValue array_v, ExcelValue k_v);
static ExcelValue left(ExcelValue string_v, ExcelValue number_of_characters_v);
static ExcelValue left_1(ExcelValue string_v);
static ExcelValue len(ExcelValue string_v);
static ExcelValue excel_log(ExcelValue number);
static ExcelValue excel_log_2(ExcelValue number, ExcelValue base);
static ExcelValue ln(ExcelValue number);
static ExcelValue excel_exp(ExcelValue number);
static ExcelValue max(int number_of_arguments, ExcelValue *arguments);
static ExcelValue min(int number_of_arguments, ExcelValue *arguments);
static ExcelValue mmult(ExcelValue a_v, ExcelValue b_v);
static ExcelValue mod(ExcelValue a_v, ExcelValue b_v);
static ExcelValue na();
static ExcelValue negative(ExcelValue a_v);
static ExcelValue excel_not(ExcelValue a_v);
static ExcelValue number_or_zero(ExcelValue maybe_number_v);
static ExcelValue npv(ExcelValue rate, int number_of_arguments, ExcelValue *arguments);
static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v);
static ExcelValue pmt_4(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v, ExcelValue final_value_v);
static ExcelValue pmt_5(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v, ExcelValue final_value_v, ExcelValue type_v);
static ExcelValue power(ExcelValue a_v, ExcelValue b_v);
static ExcelValue pv_3(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v);
static ExcelValue pv_4(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v, ExcelValue d_v);
static ExcelValue pv_5(ExcelValue a_v, ExcelValue b_v, ExcelValue c_v, ExcelValue d_v, ExcelValue e_v);
static ExcelValue excel_round(ExcelValue number_v, ExcelValue decimal_places_v);
static ExcelValue rank(ExcelValue number_v, ExcelValue range_v, ExcelValue order_v);
static ExcelValue rank_2(ExcelValue number_v, ExcelValue range_v);
static ExcelValue right(ExcelValue string_v, ExcelValue number_of_characters_v);
static ExcelValue right_1(ExcelValue string_v);
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
static ExcelValue value(ExcelValue string_v);
static ExcelValue vlookup_3(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v);
static ExcelValue vlookup(ExcelValue lookup_value_v,ExcelValue lookup_table_v, ExcelValue column_number_v, ExcelValue match_type_v);
static ExcelValue scurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration);
static ExcelValue scurve(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration, ExcelValue startYear);
static ExcelValue halfscurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration);
static ExcelValue halfscurve(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration, ExcelValue startYear);
static ExcelValue lcurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration);
static ExcelValue lcurve(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration, ExcelValue startYear);
static ExcelValue curve_5(ExcelValue curveType, ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration);
static ExcelValue curve(ExcelValue curveType, ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration, ExcelValue startYear);

static ExcelValue product(int number_of_arguments, ExcelValue *arguments);
static ExcelValue excel_floor(ExcelValue number_v, ExcelValue multiple_v);
static ExcelValue rate(ExcelValue a1, ExcelValue a2, ExcelValue a3, ExcelValue a4);
static ExcelValue excel_sqrt(ExcelValue number_v);

// My little heap for keeping pointers to memory that I need to reclaim
void **memory_that_needs_to_be_freed;
int memory_that_needs_to_be_freed_counter = 0;
int memory_that_needs_to_be_freed_size = -1;

static void free_later(void *pointer) {
	if(memory_that_needs_to_be_freed_counter >= memory_that_needs_to_be_freed_size) {
    if(memory_that_needs_to_be_freed_size <= 0) {
      memory_that_needs_to_be_freed = malloc(MEMORY_TO_BE_FREED_LATER_HEAP_INCREMENT*sizeof(void*));
      memory_that_needs_to_be_freed_size = MEMORY_TO_BE_FREED_LATER_HEAP_INCREMENT;
    } else {
      memory_that_needs_to_be_freed_size += MEMORY_TO_BE_FREED_LATER_HEAP_INCREMENT;
      memory_that_needs_to_be_freed = realloc(memory_that_needs_to_be_freed, memory_that_needs_to_be_freed_size * sizeof(void*));
      if(!memory_that_needs_to_be_freed) {
        printf("Could not allocate new memory to memory that needs to be freed array. halting.");
        exit(-1);
      }
    }
  }
	memory_that_needs_to_be_freed[memory_that_needs_to_be_freed_counter] = pointer;
	memory_that_needs_to_be_freed_counter++;
}

static void free_all_allocated_memory() {
	int i;
	for(i = 0; i < memory_that_needs_to_be_freed_counter; i++) {
		free(memory_that_needs_to_be_freed[i]);
	}
	memory_that_needs_to_be_freed_counter = 0;
}

static int variable_set[NUMBER_OF_REFS];

// Resets all cached and malloc'd values
void reset() {
  free_all_allocated_memory();
  memset(variable_set, 0, sizeof(variable_set));
}

// Handy macros

#define EXCEL_NUMBER(numberdouble) ((ExcelValue) {.type = ExcelNumber, .number = numberdouble})
#define EXCEL_STRING(stringchar) ((ExcelValue) {.type = ExcelString, .string = stringchar})
#define EXCEL_RANGE(arrayofvalues, rangerows, rangecolumns) ((ExcelValue) {.type = ExcelRange, .array = arrayofvalues, .rows = rangerows, .columns = rangecolumns})

static void * new_excel_value_array(int size) {
	ExcelValue *pointer = malloc(sizeof(ExcelValue)*size); // Freed later
	if(pointer == 0) {
		printf("Out of memory in new_excel_value_array\n");
		exit(-1);
	}
	free_later(pointer);
	return pointer;
};

// Constants
static ExcelValue ORIGINAL_EXCEL_FILENAME = {.type = ExcelString, .string = EXCEL_FILENAME };

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
		return (ExcelValue) {.type = ExcelNumber, .number = -a};
	}
}

static ExcelValue excel_char(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
  if(a <= 0) { return VALUE; }
  if(a >= 256) { return VALUE; }
  a = floor(a);
	char *string = malloc(1); // Freed later
	if(string == 0) {
	  printf("Out of memory in char");
	  exit(-1);
	}
  string[0] = a;
  free_later(string);
  return EXCEL_STRING(string);
}

static ExcelValue add(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(a + b);
}

static ExcelValue ensure_is_number(ExcelValue maybe_number_v) {
  if(maybe_number_v.type == ExcelNumber) {
    return maybe_number_v;
  }
  if(maybe_number_v.type == ExcelError) {
    return maybe_number_v;
  }
  NUMBER(maybe_number_v, maybe_number)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(maybe_number);
}

static ExcelValue number_or_zero(ExcelValue maybe_number_v) {
  if(maybe_number_v.type == ExcelNumber) {
    return maybe_number_v;
  }
  if(maybe_number_v.type == ExcelError) {
    return maybe_number_v;
  }
  return ZERO;
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

  return	EXCEL_NUMBER(log(n)/log(b));
}

static ExcelValue ln(ExcelValue number_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	NUMBER(number_v, n)
	CHECK_FOR_CONVERSION_ERROR

  if(n<=0) { return NUM; }

  return	EXCEL_NUMBER(log(n));
}

static ExcelValue excel_exp(ExcelValue number_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	NUMBER(number_v, n)
	CHECK_FOR_CONVERSION_ERROR

  return	EXCEL_NUMBER(exp(n));
}

static ExcelValue excel_sqrt(ExcelValue number_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	NUMBER(number_v, n)
	CHECK_FOR_CONVERSION_ERROR

  if(n<0) { return NUM; }

  return	EXCEL_NUMBER(sqrt(n));
}

static ExcelValue excel_floor(ExcelValue number_v, ExcelValue multiple_v) {
  CHECK_FOR_PASSED_ERROR(number_v)
  CHECK_FOR_PASSED_ERROR(multiple_v)
	NUMBER(number_v, n)
  NUMBER(multiple_v, m)
	CHECK_FOR_CONVERSION_ERROR
  if(m == 0) { return DIV0; }
  if(m < 0) { return NUM; }
  return EXCEL_NUMBER((n - fmod(n, m)));
}

static ExcelValue excel_ceiling_math_2(ExcelValue number_v, ExcelValue multiple_v) {
  return excel_ceiling_math(number_v, multiple_v, ZERO);
}

static ExcelValue excel_ceiling_math(ExcelValue number_v, ExcelValue multiple_v, ExcelValue mode_v) {
  CHECK_FOR_PASSED_ERROR(number_v)
  CHECK_FOR_PASSED_ERROR(multiple_v)
  CHECK_FOR_PASSED_ERROR(mode_v)
	NUMBER(number_v, n)
  NUMBER(multiple_v, m)
  NUMBER(mode_v, d)
	CHECK_FOR_CONVERSION_ERROR
  if(m == 0) { return ZERO; }
    
  if(d == 0 || n > 0 ) {
    double remainder = fmod(n, m);
    if(remainder == 0) {
      return number_v;
    } else {
      return EXCEL_NUMBER(((double) ceil(n/m))*m);
    }
    
    
  } else { // Need to round negative away from zero
    return negative(excel_ceiling_math(excel_abs(number_v), multiple_v, mode_v));
  }
  
  return EXCEL_NUMBER((n - fmod(n, m)));
}

static ExcelValue rate(ExcelValue periods_v, ExcelValue payment_v, ExcelValue presentValue_v, ExcelValue finalValue_v) {
  CHECK_FOR_PASSED_ERROR(periods_v)
  CHECK_FOR_PASSED_ERROR(payment_v)
  CHECK_FOR_PASSED_ERROR(presentValue_v)
  CHECK_FOR_PASSED_ERROR(finalValue_v)

  NUMBER(periods_v, periods)
  NUMBER(payment_v, payment)
  NUMBER(presentValue_v, presentValue)
  NUMBER(finalValue_v, finalValue)

  // FIXME: Only implemented the case where payment is zero
  if(payment != 0) {
    return NA;
  }

  return EXCEL_NUMBER(pow((finalValue/(-presentValue)),(1.0/periods))-1.0);
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

static ExcelValue excel_or(int array_size, ExcelValue *array) {
	int i;
	ExcelValue current_excel_value, array_result;

	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	case ExcelNumber:
		  case ExcelBoolean:
			  if(current_excel_value.number == true) return TRUE;
			  break;
		  case ExcelRange:
		  	array_result = excel_or( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
        if(array_result.type == ExcelError) return array_result;
        if(array_result.type == ExcelBoolean && array_result.number == true) return TRUE;
        break;
		  case ExcelString:
		  case ExcelEmpty:
        break;
		  case ExcelError:
			 return current_excel_value;
			 break;
		 }
	 }
	 return FALSE;
}

static ExcelValue excel_not(ExcelValue boolean_v) {
  switch (boolean_v.type) {
    case ExcelNumber:
      if(boolean_v.number == 0) return TRUE;
      return FALSE;

    case ExcelBoolean:
      if(boolean_v.number == false) return TRUE;
      return FALSE;

    case ExcelRange:
      return VALUE;

    case ExcelString:
      return VALUE;

    case ExcelEmpty:
      return TRUE;

    case ExcelError:
      return boolean_v;
  }
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
	return EXCEL_NUMBER(r.sum/r.count);
}

static ExcelValue forecast(ExcelValue required_x_v, ExcelValue known_y, ExcelValue known_x) {
  CHECK_FOR_PASSED_ERROR(required_x_v)

	NUMBER(required_x_v, required_x)
	CHECK_FOR_CONVERSION_ERROR

  if(known_x.type != ExcelRange) { return NA; }
  if(known_y.type != ExcelRange) { return NA; }

  int known_x_size = known_x.rows * known_x.columns;
  int known_y_size = known_y.rows * known_y.columns;

  int i;
  ExcelValue *x_array, *y_array;
  ExcelValue vx, vy;

  x_array = known_x.array;
  y_array = known_y.array;

  for(i=0; i<known_x_size; i++) {
    vx = x_array[i];
    if(vx.type == ExcelError) {
      return vx;
    }
  }

  for(i=0; i<known_x_size; i++) {
    vy = y_array[i];
    if(vy.type == ExcelError) {
      return vy;
    }
  }

  if(known_x_size != known_y_size) { return NA; }
  if(known_x_size == 0) { return NA; }

  ExcelValue mean_y = average(1, &known_y);
  ExcelValue mean_x = average(1, &known_x);

  if(mean_y.type == ExcelError) { return mean_y; }
  if(mean_x.type == ExcelError) { return mean_x; }

  float mx = mean_x.number;
  float my = mean_y.number;

  float b_numerator, b_denominator, b, a;

  b_denominator = 0;
  b_numerator = 0;

  for(i=0; i<known_x_size; i++) {
    vx = x_array[i];
    vy = y_array[i];
    if(vx.type != ExcelNumber) { continue; }
    if(vy.type != ExcelNumber) { continue; }

    b_denominator = b_denominator + pow(vx.number - mx, 2);
    b_numerator = b_numerator + ((vx.number - mx)*(vy.number-my));
  }

  if(b_denominator == 0) { return DIV0; }

  b = b_numerator / b_denominator;
  a = mean_y.number - (b*mean_x.number);

  return EXCEL_NUMBER(a + (b*required_x));
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
	 return EXCEL_NUMBER(n);
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
	 return EXCEL_NUMBER(n);
}

static ExcelValue divide(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return EXCEL_NUMBER(a / b);
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

static ExcelValue excel_isblank(ExcelValue value) {
  if(value.type == ExcelEmpty) {
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

  if(row_number == 0 && rows == 1) row_number = 1;
  if(column_number == 0 && columns == 1) column_number = 1;

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
		return EXCEL_RANGE(result,rows,1);
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
		return EXCEL_RANGE(result,1,columns);
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
  if(range_v.type != ExcelRange) { return VALUE; }

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
    if(x_v.type == ExcelError) { free(sorted); return x_v; };
    if(x_v.type == ExcelNumber) {
      sorted[sorted_size] = x_v.number;
      sorted_size++;
    }
  }
  // Check other bound
  if(k > sorted_size) { free(sorted); return NUM; }

  qsort(sorted, sorted_size, sizeof (double), compare_doubles);

  ExcelValue result = EXCEL_NUMBER(sorted[sorted_size - k]);
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
				if(excel_equal(lookup_value,x).number == true) return EXCEL_NUMBER(i+1);
			}
			return NA;
			break;
		case 1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
				if(lookup_value.type != x.type ) { continue; }
				if(more_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return EXCEL_NUMBER(i);
				}
			}
			return EXCEL_NUMBER(size);
			break;
		case -1:
			for(i = 0; i < size; i++ ) {
				x = array[i];
        if(lookup_value.type != x.type ) { continue; }
				if(less_than(x,lookup_value).number == true) {
					if(i==0) return NA;
					return EXCEL_NUMBER(i);
				}
			}
			return EXCEL_NUMBER(size-1);
			break;
	}
	return NA;
}

static ExcelValue excel_match_2(ExcelValue lookup_value, ExcelValue lookup_array ) {
	return excel_match(lookup_value, lookup_array, ONE);
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
		return EXCEL_NUMBER(result - within_text + 1);
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

  if(number_of_characters < 0) {
    return VALUE;
  }

	char *string;
	int string_must_be_freed = 0;
	switch (string_v.type) {
  	case ExcelString:
  		string = string_v.string;
  		break;
  	case ExcelNumber:
		  string = malloc(20); // Freed
		  if(string == 0) {
			  printf("Out of memory in left");
			  exit(-1);
		  }
		  string_must_be_freed = 1;
		  snprintf(string,20,"%0.0f",string_v.number);
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

  if(number_of_characters > strlen(string)) {
    number_of_characters = strlen(string);
  }

	char *left_string = malloc(number_of_characters+1); // Freed
	if(left_string == 0) {
	  printf("Out of memoryn in left");
	  exit(-1);
	}
	memcpy(left_string,string,number_of_characters);
	left_string[number_of_characters] = '\0';
	if(string_must_be_freed == 1) {
		free(string);
	}
	free_later(left_string);
	return EXCEL_STRING(left_string);
}

static ExcelValue left_1(ExcelValue string_v) {
	return left(string_v, ONE);
}

static ExcelValue len(ExcelValue string_v) {
	CHECK_FOR_PASSED_ERROR(string_v)
	if(string_v.type == ExcelEmpty) return ZERO;

	char *string;
	int string_must_be_freed = 0;
	switch (string_v.type) {
  	case ExcelString:
  		string = string_v.string;
  		break;
  	case ExcelNumber:
		  string = malloc(20); // Freed
		  if(string == 0) {
			  printf("Out of memory in len");
			  exit(-1);
		  }
		  snprintf(string,20,"%0.0f",string_v.number);
		  string_must_be_freed = 1;
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

  int length = strlen(string);
	if(string_must_be_freed == 1) {
		free(string);
	}
	return EXCEL_NUMBER(length);
}

static ExcelValue right(ExcelValue string_v, ExcelValue number_of_characters_v) {
	CHECK_FOR_PASSED_ERROR(string_v)
	CHECK_FOR_PASSED_ERROR(number_of_characters_v)
	if(string_v.type == ExcelEmpty) return BLANK;
	if(number_of_characters_v.type == ExcelEmpty) return BLANK;

	int number_of_characters = (int) number_from(number_of_characters_v);
	CHECK_FOR_CONVERSION_ERROR

  if(number_of_characters < 0) {
    return VALUE;
  }

	char *string;
	int string_must_be_freed = 0;
	switch (string_v.type) {
  	case ExcelString:
  		string = string_v.string;
  		break;
  	case ExcelNumber:
		  string = malloc(20); // Freed
		  if(string == 0) {
			  printf("Out of memory in right");
			  exit(-1);
		  }
		  string_must_be_freed = 1;
		  snprintf(string,20,"%0.0f",string_v.number);
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

	char *right_string = malloc(number_of_characters+1); // Freed
	if(right_string == 0) {
	  printf("Out of memory in right");
	  exit(-1);
	}
  int length = strlen(string);
  if(number_of_characters > length) {
    number_of_characters = length;
  }
  memcpy(right_string,string+length-number_of_characters,number_of_characters);
  right_string[number_of_characters] = '\0';
  if(string_must_be_freed == 1) {
    free(string);
  }
  free_later(right_string);
  return EXCEL_STRING(right_string);
}

static ExcelValue right_1(ExcelValue string_v) {
	return right(string_v, ONE);
}

static ExcelValue iferror(ExcelValue value, ExcelValue value_if_error) {
	if(value.type == ExcelError) return value_if_error;
	return value;
}

static ExcelValue iserr(ExcelValue value) {
	if(value.type == ExcelError) {
    if(value.number == NA.number) {
      return FALSE;
    } else {
      return TRUE;
    }
  } else {
    return FALSE;
  }
}

static ExcelValue iserror(ExcelValue value) {
	if(value.type == ExcelError) {
    return TRUE;
  } else {
    return FALSE;
  }
}



// Order is TRUE, FALSE, String, Number; Blank is zero
static ExcelValue more_than(ExcelValue a_v, ExcelValue b_v) {
  CHECK_FOR_PASSED_ERROR(a_v)
  CHECK_FOR_PASSED_ERROR(b_v)

  if(a_v.type == ExcelEmpty) { a_v = ZERO; }
  if(b_v.type == ExcelEmpty) { b_v = ZERO; }

  switch (a_v.type) {
    case ExcelString:
      switch (b_v.type) {
        case ExcelString:
          if(strcasecmp(a_v.string,b_v.string) <= 0 ) {return FALSE;} else {return TRUE;}
        case ExcelNumber:
          return TRUE;
        case ExcelBoolean:
          return FALSE;
        // Following shouldn't happen
        case ExcelEmpty:
        case ExcelError:
        case ExcelRange:
          return NA;
      }
    case ExcelBoolean:
      switch (b_v.type) {
        case ExcelBoolean:
          if(a_v.number == true) {
            if (b_v.number == true) { return FALSE; } else { return TRUE; }
          } else { // a_v == FALSE
            return FALSE;
          }
        case ExcelString:
        case ExcelNumber:
          return TRUE;
        // Following shouldn't happen
        case ExcelEmpty:
        case ExcelError:
        case ExcelRange:
          return NA;
      }
    case ExcelNumber:
      switch (b_v.type) {
        case ExcelNumber:
          if(a_v.number > b_v.number) { return TRUE; } else { return FALSE; }
        case ExcelString:
        case ExcelBoolean:
          return FALSE;
        // Following shouldn't happen
        case ExcelEmpty:
        case ExcelError:
        case ExcelRange:
          return NA;
      }
    // Following shouldn't happen
    case ExcelEmpty:
    case ExcelError:
    case ExcelRange:
      return NA;
  }
  // Shouldn't reach here
  return NA;
}

static ExcelValue more_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
  ExcelValue opposite = less_than(a_v, b_v);
  switch (opposite.type) {
    case ExcelBoolean:
      if(opposite.number == true) { return FALSE; } else { return TRUE; }
    case ExcelError:
      return opposite;
    // Shouldn't reach below
    case ExcelNumber:
    case ExcelString:
    case ExcelEmpty:
    case ExcelRange:
      return NA;
  }
}

// Order is TRUE, FALSE, String, Number; Blank is zero
static ExcelValue less_than(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

  if(a_v.type == ExcelEmpty) { a_v = ZERO; }
  if(b_v.type == ExcelEmpty) { b_v = ZERO; }

	switch (a_v.type) {
    case ExcelString:
      switch (b_v.type) {
        case ExcelString:
          if(strcasecmp(a_v.string, b_v.string) >= 0 )  {
            return FALSE;
          } else {
            return TRUE;
          }
        case ExcelNumber:
          return FALSE;
        case ExcelBoolean:
          return TRUE;
        // The following shouldn't happen
        // FIXME: Should abort if it does
        case ExcelError:
        case ExcelRange:
        case ExcelEmpty:
          return NA;
      }
  	case ExcelNumber:
      switch(b_v.type) {
        case ExcelNumber:
          if(a_v.number < b_v.number) {
            return TRUE;
          } else {
            return FALSE;
          }
        case ExcelBoolean:
        case ExcelString:
          return TRUE;
        // The following shouldn't happen
        // FIXME: Should abort if it does
        case ExcelError:
        case ExcelRange:
        case ExcelEmpty:
          return NA;
      }
    case ExcelBoolean:
      switch(b_v.type) {
        case ExcelBoolean:
          if(a_v.number == true) {
            return FALSE;
          } else { // a_v.number == false
            if(b_v.number == true) {return TRUE;} else {return FALSE;}
          }
        case ExcelString:
        case ExcelNumber:
          return FALSE;
        // The following shouldn't happen
        // FIXME: Should abort if it does
        case ExcelError:
        case ExcelRange:
        case ExcelEmpty:
          return NA;
      }
    // The following shouldn't happen
    // FIXME: Should abort if it does
    case ExcelError:
    case ExcelRange:
    case ExcelEmpty:
      return VALUE;
  }
  // Shouldn't reach here
  return NA;
}

static ExcelValue less_than_or_equal(ExcelValue a_v, ExcelValue b_v) {
  ExcelValue opposite = more_than(a_v, b_v);
  switch (opposite.type) {
    case ExcelBoolean:
      if(opposite.number == true) { return FALSE; } else { return TRUE; }
    case ExcelError:
      return opposite;
    // Shouldn't reach below
    case ExcelNumber:
    case ExcelString:
    case ExcelEmpty:
    case ExcelRange:
      return VALUE;
  }
}

static ExcelValue subtract(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(a - b);
}

static ExcelValue multiply(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(a * b);
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
	return EXCEL_NUMBER(total);
}

static ExcelValue npv(ExcelValue rate_v, int number_of_arguments, ExcelValue *arguments) {
	CHECK_FOR_PASSED_ERROR(rate_v)
	NUMBER(rate_v, rate)
	CHECK_FOR_CONVERSION_ERROR
  if(rate == -1) { return DIV0; }

  double npv = 0;
  int n = 1;
  int i;
  int j;
  double v;
  ExcelValue r;
  ExcelValue r2;
  ExcelValue *range;

  for(i=0;i<number_of_arguments;i++) {
    r = arguments[i];
    if(r.type == ExcelError) { return r; }
    if(r.type == ExcelRange) {
      range = r.array;
      for(j=0;j<(r.columns*r.rows);j++) {
        r2 = range[j];
        if(r2.type == ExcelError) { return r2; }
        v = number_from(r2);
        if(conversion_error) { conversion_error = 0; return VALUE; }
        npv = npv + (v/pow(1+rate, n));
        n++;
      }
    } else {
      v = number_from(r);
      if(conversion_error) { conversion_error = 0; return VALUE; }
      npv = npv + (v/pow(1+rate, n));
      n++;
    }
  }
  return EXCEL_NUMBER(npv);
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
	return EXCEL_NUMBER(biggest_number_found);
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
	return EXCEL_NUMBER(smallest_number_found);
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
  return EXCEL_RANGE(result, rows, columns);
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
      result[(i*b_columns)+j] = EXCEL_NUMBER(sum);
    }
  }
  return EXCEL_RANGE(result, a_rows, b_columns);
}

static ExcelValue mod(ExcelValue a_v, ExcelValue b_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	CHECK_FOR_PASSED_ERROR(b_v)

	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return EXCEL_NUMBER(fmod(a,b));
}

static ExcelValue na() {
  return NA;
}

static ExcelValue negative(ExcelValue a_v) {
	CHECK_FOR_PASSED_ERROR(a_v)
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(-a);
}

static ExcelValue pmt(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v) {
	CHECK_FOR_PASSED_ERROR(rate_v)
	CHECK_FOR_PASSED_ERROR(number_of_periods_v)
	CHECK_FOR_PASSED_ERROR(present_value_v)

	NUMBER(rate_v,rate)
	NUMBER(number_of_periods_v,number_of_periods)
	NUMBER(present_value_v,present_value)
	CHECK_FOR_CONVERSION_ERROR

	if(rate == 0) return EXCEL_NUMBER(-(present_value / number_of_periods));
	return EXCEL_NUMBER(-present_value*(rate*(pow((1+rate),number_of_periods)))/((pow((1+rate),number_of_periods))-1));
}

static ExcelValue pmt_4(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v, ExcelValue final_value_v) {
  CHECK_FOR_PASSED_ERROR(final_value_v)

    NUMBER(final_value_v, final_value)
    CHECK_FOR_CONVERSION_ERROR

    if(final_value == 0) return pmt(rate_v, number_of_periods_v, present_value_v);
    printf("PMT with non-zero final_value not implemented. halting.");
    exit(-1);
}

static ExcelValue pmt_5(ExcelValue rate_v, ExcelValue number_of_periods_v, ExcelValue present_value_v, ExcelValue final_value_v, ExcelValue type_v) {
  CHECK_FOR_PASSED_ERROR(type_v)

    NUMBER(type_v, type)
    CHECK_FOR_CONVERSION_ERROR

    if(type == 0) return pmt(rate_v, number_of_periods_v, present_value_v);
    printf("PMT with non-zero type not implemented. halting.");
    exit(-1);
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

  return EXCEL_NUMBER(present_value);
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
    return EXCEL_NUMBER(result);
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
  return EXCEL_NUMBER(ranked);
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

	return EXCEL_NUMBER( round(number * multiple) / multiple );
}

static ExcelValue rounddown(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)

	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR

	double multiple = pow(10,decimal_places);

	return EXCEL_NUMBER( trunc(number * multiple) / multiple );
}

static ExcelValue roundup(ExcelValue number_v, ExcelValue decimal_places_v) {
	CHECK_FOR_PASSED_ERROR(number_v)
	CHECK_FOR_PASSED_ERROR(decimal_places_v)

	NUMBER(number_v, number)
	NUMBER(decimal_places_v, decimal_places)
	CHECK_FOR_CONVERSION_ERROR

	double multiple = pow(10,decimal_places);
	if(number < 0) return EXCEL_NUMBER( floor(number * multiple) / multiple );
	return EXCEL_NUMBER( ceil(number * multiple) / multiple );
}

static ExcelValue excel_int(ExcelValue number_v) {
	CHECK_FOR_PASSED_ERROR(number_v)

	NUMBER(number_v, number)
	CHECK_FOR_CONVERSION_ERROR

	return EXCEL_NUMBER(floor(number));
}

static ExcelValue string_join(int number_of_arguments, ExcelValue *arguments) {
	int allocated_length = 100;
	int used_length = 0;
	char *string = malloc(allocated_length); // Freed later
	if(string == 0) {
	  printf("Out of memory in string_join");
	  exit(-1);
	}
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
		  	  printf("Out of memory in string join");
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
        free(string);
        return current_v;
	  	case ExcelRange:
        free(string);
        return VALUE;
		}
		current_string_length = strlen(current_string);
		if( (used_length + current_string_length + 1) > allocated_length) {
			allocated_length = used_length + current_string_length + 1 + 100;
			string = realloc(string,allocated_length);
      if(!string) {
        printf("Out of memory in string join realloc trying to increase to %d", allocated_length);
        exit(-1);
      }
		}
		memcpy(string + used_length, current_string, current_string_length);
		if(must_free_current_string == 1) {
			free(current_string);
		}
		used_length = used_length + current_string_length;
	} // Finished looping through passed strings
	string = realloc(string,used_length+1);
  if(!string) {
    printf("Out of memory in string join realloc trying to increase to %d", used_length+1);
    exit(-1);
  }
  string[used_length] = '\0';
	free_later(string);
	return EXCEL_STRING(string);
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


// FIXME: Check if this deals with errors correctly
static ExcelValue filter_range(ExcelValue original_range_v, int number_of_arguments, ExcelValue *arguments) {
  // First, set up the original_range
  //CHECK_FOR_PASSED_ERROR(original_range_v);

  // Set up the sum range
  ExcelValue *original_range;
  int original_range_rows, original_range_columns;

  if(original_range_v.type == ExcelRange) {
    original_range = original_range_v.array;
    original_range_rows = original_range_v.rows;
    original_range_columns = original_range_v.columns;
  } else {
    original_range = (ExcelValue*) new_excel_value_array(1);
	  original_range[0] = original_range_v;
    original_range_rows = 1;
    original_range_columns = 1;
  }

  // This is the filtered range
  ExcelValue *filtered_range = new_excel_value_array(original_range_rows*original_range_columns);
  int number_of_filtered_values = 0;

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
      if(current_value.rows != original_range_rows) return VALUE;
      if(current_value.columns != original_range_columns) return VALUE;
    } else {
      if(original_range_rows != 1) return VALUE;
      if(original_range_columns != 1) return VALUE;
      ExcelValue *tmp_array2 =  (ExcelValue*) new_excel_value_array(1);
      tmp_array2[0] = current_value;
      criteria_range[i] =  EXCEL_RANGE(tmp_array2,1,1);
    }
  }

  // Now go through and set up the criteria
  ExcelComparison *criteria =  malloc(sizeof(ExcelComparison)*number_of_criteria); // freed at end of function
  if(criteria == 0) {
	  printf("Out of memory in filter_range\n");
	  exit(-1);
  }
  char *s;
  char *new_comparator;

  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[(i*2)+1];

    if(current_value.type == ExcelString) {
      s = current_value.string;
      if(s[0] == '<') {
        if( s[1] == '>') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = NotEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else if(s[1] == '=') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = LessThanOrEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else {
          new_comparator = strndup(s+1,strlen(s)-1);
          free_later(new_comparator);
          criteria[i].type = LessThan;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        }
      } else if(s[0] == '>') {
        if(s[1] == '=') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = MoreThanOrEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else {
          new_comparator = strndup(s+1,strlen(s)-1);
          free_later(new_comparator);
          criteria[i].type = MoreThan;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        }
      } else if(s[0] == '=') {
        new_comparator = strndup(s+1,strlen(s)-1);
        free_later(new_comparator);
        criteria[i].type = Equal;
        criteria[i].comparator = EXCEL_STRING(new_comparator);
      } else {
        criteria[i].type = Equal;
        criteria[i].comparator = current_value;
      }
    } else {
      criteria[i].type = Equal;
      criteria[i].comparator = current_value;
    }
  }

  int size = original_range_columns * original_range_rows;
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

      // For the purposes of comparison, treates a blank criteria as matching zeros.
      if(comparator.type == ExcelEmpty) {
        comparator = ZERO;
      }

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

            // Special case, empty strings don't match zeros here
            if(strlen(value_to_be_checked.string) == 0) {
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
          free(criteria);
          return VALUE;
      }
      if(passed == 0) break;
    }
    if(passed == 1) {
      current_value = original_range[j];
      if(current_value.type == ExcelError) {
        free(criteria);
        return current_value;
      } else if(current_value.type == ExcelNumber) {
        filtered_range[number_of_filtered_values] = current_value;
        number_of_filtered_values += 1;
      }
    }
  }
  // Tidy up
  free(criteria);
  return EXCEL_RANGE(filtered_range, number_of_filtered_values, 1);
}

static ExcelValue sumifs(ExcelValue sum_range_v, int number_of_arguments, ExcelValue *arguments) {
  ExcelValue filtered_range = filter_range(sum_range_v, number_of_arguments, arguments);
  return sum(1,&filtered_range);
}

static ExcelValue countifs(int number_of_arguments, ExcelValue *arguments) {
  if(number_of_arguments < 2) { return NA;}
  // Set up the sum range
  ExcelValue range = arguments[0];
  int rows, columns;

  if(range.type == ExcelRange) {
    rows = range.rows;
    columns = range.columns;
  } else {
    rows = 1;
    columns = 1;
  }

  int count = 0;

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
      if(current_value.rows != rows) return VALUE;
      if(current_value.columns != columns) return VALUE;
    } else {
      if(rows != 1) return VALUE;
      if(columns != 1) return VALUE;
      ExcelValue *tmp_array2 =  (ExcelValue*) new_excel_value_array(1);
      tmp_array2[0] = current_value;
      criteria_range[i] = EXCEL_RANGE(tmp_array2,1,1);
    }
  }

  // Now go through and set up the criteria
  ExcelComparison *criteria =  malloc(sizeof(ExcelComparison)*number_of_criteria); // freed at end of function
  if(criteria == 0) {
	  printf("Out of memory in filter_range\n");
	  exit(-1);
  }
  char *s;
  char *new_comparator;

  for(i = 0; i < number_of_criteria; i++) {
    current_value = arguments[(i*2)+1];

    if(current_value.type == ExcelString) {
      s = current_value.string;
      if(s[0] == '<') {
        if( s[1] == '>') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = NotEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else if(s[1] == '=') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = LessThanOrEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else {
          new_comparator = strndup(s+1,strlen(s)-1);
          free_later(new_comparator);
          criteria[i].type = LessThan;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        }
      } else if(s[0] == '>') {
        if(s[1] == '=') {
          new_comparator = strndup(s+2,strlen(s)-2);
          free_later(new_comparator);
          criteria[i].type = MoreThanOrEqual;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        } else {
          new_comparator = strndup(s+1,strlen(s)-1);
          free_later(new_comparator);
          criteria[i].type = MoreThan;
          criteria[i].comparator = EXCEL_STRING(new_comparator);
        }
      } else if(s[0] == '=') {
        new_comparator = strndup(s+1,strlen(s)-1);
        free_later(new_comparator);
        criteria[i].type = Equal;
        criteria[i].comparator = EXCEL_STRING(new_comparator);
      } else {
        criteria[i].type = Equal;
        criteria[i].comparator = current_value;
      }
    } else {
      criteria[i].type = Equal;
      criteria[i].comparator = current_value;
    }
  }

  int size = columns * rows;
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

      // For the purposes of comparison, treates a blank criteria as matching zeros.
      if(comparator.type == ExcelEmpty) {
        comparator = ZERO;
      }

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

            // Special case, empty strings don't match zeros here
            if(strlen(value_to_be_checked.string) == 0) {
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
          free(criteria);
          return VALUE;
      }
      if(passed == 0) break;
    }
    if(passed == 1) {
        count += 1;
    }
  }
  // Tidy up
  free(criteria);
  return EXCEL_NUMBER(count);
}

static ExcelValue averageifs(ExcelValue average_range_v, int number_of_arguments, ExcelValue *arguments) {
  ExcelValue filtered_range = filter_range(average_range_v, number_of_arguments, arguments);
  return average(1,&filtered_range);
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
	  printf("Out of memory in sumproduct\n");
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
        if(current_value.rows != rows || current_value.columns != columns) { free(ranges);  return VALUE; }
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
        if(rows != 1 && columns !=1) { free(ranges); return VALUE; }
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
  return EXCEL_NUMBER(sum);
}

static ExcelValue product(int number_of_arguments, ExcelValue *arguments) {
  if(number_of_arguments <1) return VALUE;

  int a,b;
  ExcelValue sub_total;
  ExcelValue current_value;
  int sub_total_array_size;
  ExcelValue *sub_total_array;
  ExcelValue sub_total_value;
  double total = 0;

  // Extract arrays from each of the given ranges, checking for errors and bounds as we go
  for(a=0;a<number_of_arguments;a++) {
    current_value = arguments[a];
    switch(current_value.type) {
      case ExcelRange:
        sub_total_array_size = current_value.rows * current_value.columns;
        sub_total_array = current_value.array;
        // We don't use recursion, because we need to check if
        // the result is 0 becaues a zero, or zero because all blank.
        for(b=0;b<sub_total_array_size;b++) {
          sub_total_value = sub_total_array[b];
          switch(sub_total_value.type) {
            case ExcelError:
              return sub_total_value;
              break;

            case ExcelNumber:
              // We do this rather than starting with total = 1
              // so that the product of all blanks is zero
              if(total == 0) {
                total = sub_total_value.number;
              } else {
                total *= sub_total_value.number;
              }
              break;

            default:
              // Skip
              break;
          }
        }
        break;

      case ExcelError:
        return current_value;
        break;

      case ExcelNumber:
        if(total == 0) {
          total = current_value.number;
        } else {
          total *= current_value.number;
        }
        break;

      default:
        // Skip
        break;
    }
  }

  return EXCEL_NUMBER(total);
}

// FIXME: This could do with being done properly, rather than
// on a case by case basis.
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
    return EXCEL_STRING("");
  }

  if(format_v.type == ExcelNumber && format_v.number == 0) {
    format_v = EXCEL_STRING("0");
  }

  if(number_v.type == ExcelString) {
 	 	s = number_v.string;
		if (s == NULL || *s == '\0' || isspace(*s)) {
			number_v = ZERO;
		}
		n = strtod (s, &p);
		if(*p == '\0') {
		  number_v = EXCEL_NUMBER(n);
		}
  }

  if(number_v.type != ExcelNumber) {
    return number_v;
  }

  if(format_v.type != ExcelString) {
    return format_v;
  }

  // FIXME: Too little?
  s = malloc(100);
  setlocale(LC_ALL,"");

  if(strcmp(format_v.string,"0%") == 0) {
    snprintf(s, 99, "%0.0f%%", number_v.number*100);
  } else if(strcmp(format_v.string,"0.0%") == 0) {
    snprintf(s, 99, "%0.1f%%", number_v.number*100);
  } else if(strcmp(format_v.string,"0") == 0) {
    snprintf(s, 99, "%0.0f",number_v.number);
  } else if(strcmp(format_v.string,"0.0") == 0) {
    snprintf(s, 99, "%0.1f",number_v.number);
  } else if(strcmp(format_v.string,"0.00") == 0) {
    snprintf(s, 99, "%0.2f",number_v.number);
  } else if(strcmp(format_v.string,"0.000") == 0) {
    snprintf(s, 99, "%0.3f",number_v.number);
  } else if(strcmp(format_v.string,"#,##") == 0) {
    snprintf(s, 99, "%'0.0f",number_v.number);
  } else if(strcmp(format_v.string,"#,##0") == 0) {
    snprintf(s, 99, "%'0.0f",number_v.number);
  } else if(strcmp(format_v.string,"#,##0.0") == 0) {
    snprintf(s, 99, "%'0.1f",number_v.number);
  } else if(strcmp(format_v.string,"#,##0.00") == 0) {
    snprintf(s, 99, "%'0.2f",number_v.number);
  } else if(strcmp(format_v.string,"#,##0.000") == 0) {
    snprintf(s, 99, "%'0.3f",number_v.number);
  } else if(strcmp(format_v.string,"0000") == 0) {
    snprintf(s, 99, "%04.0f",number_v.number);
  } else {
    snprintf(s, 99, "Text format not recognised");
  }

  free_later(s);
  result = EXCEL_STRING(s);
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


static ExcelValue value(ExcelValue string_v) {
	CHECK_FOR_PASSED_ERROR(string_v)
	NUMBER(string_v, a)
	CHECK_FOR_CONVERSION_ERROR
	return EXCEL_NUMBER(a);
}

static ExcelValue scurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration) {
  ExcelValue startYear = EXCEL_NUMBER(2018);
  return scurve(currentYear, startValue, endValue, duration, startYear);
}

static ExcelValue halfscurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration) {
  ExcelValue startYear = EXCEL_NUMBER(2018);
  return halfscurve(currentYear, startValue, endValue, duration, startYear);
}

static ExcelValue lcurve_4(ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration) {
  ExcelValue startYear = EXCEL_NUMBER(2018);
  return lcurve(currentYear, startValue, endValue, duration, startYear);
}

static ExcelValue curve_5(ExcelValue curveType, ExcelValue currentYear, ExcelValue startValue, ExcelValue endValue, ExcelValue duration) {
  ExcelValue startYear = EXCEL_NUMBER(2018);
  return curve(curveType, currentYear, startValue, endValue, duration, startYear);
}

static ExcelValue scurve(ExcelValue currentYear_v, ExcelValue startValue_v, ExcelValue endValue_v, ExcelValue duration_v, ExcelValue startYear_v) {

	NUMBER(currentYear_v, currentYear)
	NUMBER(startValue_v, startValue)
	NUMBER(endValue_v, endValue)
	NUMBER(duration_v, duration)
	NUMBER(startYear_v, startYear)
	CHECK_FOR_CONVERSION_ERROR

  if(currentYear < startYear) {
    return startValue_v;
  }
  double x = (currentYear - startYear) / duration;
  double x0 = 0.0;
  double a = endValue - startValue;
  double sc = 0.999;
  double eps = 1.0 - sc;
  double mu = 0.5;
  double beta = (mu - 1.0) / log(1.0 / sc - 1);
  double scurve = a * (pow((exp(-(x - mu) / beta) + 1),-1) - pow((exp(-(x0 - mu) / beta) + 1),-1)) + startValue;
  return EXCEL_NUMBER(scurve);
}

static ExcelValue halfscurve(ExcelValue currentYear_v, ExcelValue startValue_v, ExcelValue endValue_v, ExcelValue duration_v, ExcelValue startYear_v) {

	NUMBER(currentYear_v, currentYear)
	NUMBER(startValue_v, startValue)
	NUMBER(endValue_v, endValue)
	NUMBER(duration_v, duration)
	NUMBER(startYear_v, startYear)
	CHECK_FOR_CONVERSION_ERROR

  if(currentYear < startYear) {
    return startValue_v;
  }

  ExcelValue newCurrentYear = EXCEL_NUMBER(currentYear + duration);
  ExcelValue newDuration = EXCEL_NUMBER(duration *2);
  ExcelValue result_v = scurve(newCurrentYear, startValue_v, endValue_v, newDuration, startYear_v);

	NUMBER(result_v, result)
	CHECK_FOR_CONVERSION_ERROR

  return EXCEL_NUMBER(result -((endValue - startValue)/2.0));
}

static ExcelValue lcurve(ExcelValue currentYear_v, ExcelValue startValue_v, ExcelValue endValue_v, ExcelValue duration_v, ExcelValue startYear_v) {

	NUMBER(currentYear_v, currentYear)
	NUMBER(startValue_v, startValue)
	NUMBER(endValue_v, endValue)
	NUMBER(duration_v, duration)
	NUMBER(startYear_v, startYear)
	CHECK_FOR_CONVERSION_ERROR

  if(currentYear > (startYear + duration)) {
    return endValue_v;
  }

  if(currentYear < startYear) {
    return startValue_v;
  }

  double result = startValue + (((endValue - startValue) / duration) * (currentYear - startYear));
  return EXCEL_NUMBER(result);
}

static ExcelValue curve(ExcelValue type_v, ExcelValue currentYear_v, ExcelValue startValue_v, ExcelValue endValue_v, ExcelValue duration_v, ExcelValue startYear_v) {

  if(type_v.type == ExcelString && strcasecmp(type_v.string, "s") == 0 ) {
    return scurve(currentYear_v, startValue_v, endValue_v, duration_v, startYear_v);
  }

  if(type_v.type == ExcelString && strcasecmp(type_v.string, "hs") == 0 ) {
    return halfscurve(currentYear_v, startValue_v, endValue_v, duration_v, startYear_v);
  }

  return lcurve(currentYear_v, startValue_v, endValue_v, duration_v, startYear_v);
}



// Allows numbers to be 0.1% different
static ExcelValue roughly_equal(ExcelValue a_v, ExcelValue b_v) {

  if(a_v.type == ExcelEmpty && b_v.type == ExcelNumber && b_v.number == 0) return TRUE;
  if(b_v.type == ExcelEmpty && a_v.type == ExcelNumber && a_v.number == 0) return TRUE;

	if(a_v.type != b_v.type) return FALSE;

  float epsilon, difference;

	switch (a_v.type) {
  	case ExcelNumber:
      // FIXME: Arbitrary choice of epsilons
      if(b_v.number > -1e-6 && b_v.number < 1e-6) {
        epsilon = 1e-6;
      } else {
        epsilon = b_v.number * 0.001;
      }
      if(epsilon < 0) epsilon = -epsilon;
      difference = a_v.number - b_v.number;
      if(difference < 0) difference = -difference;
      if(difference <= epsilon) return TRUE;
      // For debuging: printf("a: %e b:%e d: %e e: %e", a_v.number, b_v.number, difference, epsilon);
      return FALSE;
	  case ExcelBoolean:
	  case ExcelEmpty:
			if(a_v.number != b_v.number) return FALSE;
			return TRUE;
	  case ExcelString:
	  	if(strcasecmp(a_v.string,b_v.string) != 0 ) return FALSE;
		  return TRUE;
  	case ExcelError:
			if(a_v.number != b_v.number) return FALSE;
			return TRUE;
  	case ExcelRange:
  		return NA;
  }
  return FALSE;
}


static void assert_equal(ExcelValue expected, ExcelValue actual, char location[]) {
  ExcelValue comparison = roughly_equal(actual, expected);
  if(comparison.type == ExcelBoolean && comparison.number == 1) {
    putchar('.');
  } else {
    printf("\n\nFailed at %s\n", location);
    printf("Expected: ");
    inspect_excel_value(expected);
    printf("Got:      ");
    inspect_excel_value(actual);
    putchar('\n');
  }
}
// End of the generic c functions

// Start of the file specific functions

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
ExcelValue referencing_e68();
ExcelValue referencing_f68();
ExcelValue referencing_e69();
ExcelValue referencing_f69();
ExcelValue referencing_e70();
ExcelValue referencing_f70();
ExcelValue referencing_e72();
ExcelValue referencing_f72();
ExcelValue referencing_g72();
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
ExcelValue tables_g17();
ExcelValue s_innapropriate_sheet_name__c4();
ExcelValue ranges_f1();
ExcelValue ranges_f2();
ExcelValue ranges_f3();
ExcelValue ranges_a5();
ExcelValue ranges_b5();
ExcelValue ranges_c5();
ExcelValue ranges_d5();
ExcelValue referencing_f35();
ExcelValue referencing_g35();
ExcelValue referencing_h35();
ExcelValue referencing_i35();
ExcelValue referencing_k35();
ExcelValue referencing_l35();
ExcelValue referencing_f36();
ExcelValue referencing_g36();
ExcelValue referencing_h36();
ExcelValue referencing_i36();
ExcelValue referencing_k36();
ExcelValue referencing_l36();
ExcelValue referencing_g37();
ExcelValue referencing_h37();
ExcelValue referencing_i37();
ExcelValue referencing_j37();
ExcelValue referencing_k37();
ExcelValue referencing_l37();
ExcelValue referencing_f38();
ExcelValue referencing_g38();
ExcelValue referencing_h38();
ExcelValue referencing_j38();
ExcelValue referencing_k38();
ExcelValue referencing_l38();
ExcelValue referencing_f39();
ExcelValue referencing_g39();
ExcelValue referencing_i39();
ExcelValue referencing_j39();
ExcelValue referencing_k39();
ExcelValue referencing_l39();
ExcelValue referencing_f40();
ExcelValue referencing_h40();
ExcelValue referencing_i40();
ExcelValue referencing_k40();
ExcelValue referencing_l40();
ExcelValue referencing_f41();
ExcelValue referencing_h41();
ExcelValue referencing_i41();
ExcelValue referencing_k41();
ExcelValue referencing_l41();
ExcelValue referencing_g42();
ExcelValue referencing_h42();
ExcelValue referencing_i42();
ExcelValue referencing_j42();
ExcelValue referencing_k42();
ExcelValue referencing_n42();
ExcelValue referencing_g43();
ExcelValue referencing_h43();
ExcelValue referencing_i43();
ExcelValue referencing_j43();
ExcelValue referencing_k43();
ExcelValue referencing_n43();
ExcelValue referencing_f44();
ExcelValue referencing_g44();
ExcelValue referencing_h44();
ExcelValue referencing_i44();
ExcelValue referencing_j44();
ExcelValue referencing_k44();
ExcelValue referencing_f45();
ExcelValue referencing_h45();
ExcelValue referencing_i45();
ExcelValue referencing_k45();
ExcelValue referencing_l45();
ExcelValue referencing_f46();
ExcelValue referencing_i46();
ExcelValue referencing_j46();
ExcelValue referencing_k46();
ExcelValue referencing_l46();
ExcelValue referencing_f47();
ExcelValue referencing_g47();
ExcelValue referencing_h47();
ExcelValue referencing_i47();
ExcelValue referencing_j47();
ExcelValue referencing_l47();
ExcelValue common0();
ExcelValue common1();
ExcelValue common2();
ExcelValue common3();
ExcelValue common7();
ExcelValue common9();
ExcelValue common10();
// starting the value constants
static ExcelValue constant1 = {.type = ExcelString, .string = "Hello"};
static ExcelValue constant2 = {.type = ExcelNumber, .number = 3.1415};
static ExcelValue constant3 = {.type = ExcelString, .string = "Simple"};
static ExcelValue constant4 = {.type = ExcelString, .string = "Sharing"};
static ExcelValue constant5 = {.type = ExcelNumber, .number = 267.7467614837482};
static ExcelValue constant6 = {.type = ExcelString, .string = "Shared"};
static ExcelValue constant7 = {.type = ExcelString, .string = "Array (single)"};
static ExcelValue constant8 = {.type = ExcelString, .string = "Arraying (multiple)"};
static ExcelValue constant9 = {.type = ExcelString, .string = "Not Eight"};
static ExcelValue constant10 = {.type = ExcelString, .string = "Arrayed (multiple)"};
static ExcelValue constant11 = {.type = ExcelString, .string = "This sheet"};
static ExcelValue constant12 = {.type = ExcelString, .string = "Other sheet"};
static ExcelValue constant13 = {.type = ExcelString, .string = "Standard"};
static ExcelValue constant14 = {.type = ExcelString, .string = "Column"};
static ExcelValue constant15 = {.type = ExcelString, .string = "Row"};
static ExcelValue constant16 = {.type = ExcelString, .string = "Named"};
static ExcelValue constant17 = {.type = ExcelString, .string = "Reference"};
static ExcelValue constant18 = {.type = ExcelNumber, .number = 1.4535833325868115};
static ExcelValue constant19 = {.type = ExcelNumber, .number = 1.511726665890284};
static ExcelValue constant20 = {.type = ExcelNumber, .number = 1.5407983325420203};
static ExcelValue constant21 = {.type = ExcelNumber, .number = 9.054545454545455};
static ExcelValue constant22 = {.type = ExcelNumber, .number = 12.0};
static ExcelValue constant23 = {.type = ExcelNumber, .number = 18.0};
static ExcelValue constant24 = {.type = ExcelNumber, .number = 0.3681150635671386};
static ExcelValue constant25 = {.type = ExcelNumber, .number = 0.40588480110308967};
static ExcelValue constant26 = {.type = ExcelNumber, .number = 0.42190146532760275};
static ExcelValue constant27 = {.type = ExcelNumber, .number = 0.651};
static ExcelValue constant28 = {.type = ExcelString, .string = "Technology efficiencies -- hot water -- annual mean"};
static ExcelValue constant29 = {.type = ExcelString, .string = "% of input energy"};
static ExcelValue constant30 = {.type = ExcelString, .string = "Electricity (delivered to end user)"};
static ExcelValue constant31 = {.type = ExcelString, .string = "Electricity (supplied to grid)"};
static ExcelValue constant32 = {.type = ExcelString, .string = "Solid hydrocarbons"};
static ExcelValue constant33 = {.type = ExcelString, .string = "Liquid hydrocarbons"};
static ExcelValue constant34 = {.type = ExcelString, .string = "Gaseous hydrocarbons"};
static ExcelValue constant35 = {.type = ExcelString, .string = "Heat transport"};
static ExcelValue constant36 = {.type = ExcelString, .string = "Environmental heat"};
static ExcelValue constant37 = {.type = ExcelString, .string = "Heating & cooling"};
static ExcelValue constant38 = {.type = ExcelString, .string = "Conversion losses"};
static ExcelValue constant39 = {.type = ExcelString, .string = "Balance"};
static ExcelValue constant40 = {.type = ExcelString, .string = "Code"};
static ExcelValue constant41 = {.type = ExcelString, .string = "Technology"};
static ExcelValue constant42 = {.type = ExcelString, .string = "Notes"};
static ExcelValue constant43 = {.type = ExcelString, .string = "V.01"};
static ExcelValue constant44 = {.type = ExcelString, .string = "V.02"};
static ExcelValue constant45 = {.type = ExcelString, .string = "V.03"};
static ExcelValue constant46 = {.type = ExcelString, .string = "V.04"};
static ExcelValue constant47 = {.type = ExcelString, .string = "V.05"};
static ExcelValue constant48 = {.type = ExcelString, .string = "V.07"};
static ExcelValue constant49 = {.type = ExcelString, .string = "R.07"};
static ExcelValue constant50 = {.type = ExcelString, .string = "H.01"};
static ExcelValue constant51 = {.type = ExcelString, .string = "X.01"};
static ExcelValue constant52 = {.type = ExcelString, .string = "Gas boiler (old)"};
static ExcelValue constant53 = {.type = ExcelNumber, .number = -1.0};
static ExcelValue constant54 = {.type = ExcelNumber, .number = 0.76};
static ExcelValue constant55 = {.type = ExcelNumber, .number = 0.24};
static ExcelValue constant56 = {.type = ExcelString, .string = "Gas boiler (new)"};
static ExcelValue constant57 = {.type = ExcelNumber, .number = 0.91};
static ExcelValue constant58 = {.type = ExcelNumber, .number = 0.09};
static ExcelValue constant59 = {.type = ExcelString, .string = "Resistive heating"};
static ExcelValue constant60 = {.type = ExcelString, .string = "Oil-fired boiler"};
static ExcelValue constant61 = {.type = ExcelNumber, .number = 0.97};
static ExcelValue constant62 = {.type = ExcelNumber, .number = 0.03};
static ExcelValue constant63 = {.type = ExcelNumber, .number = -2.7755575615628914e-17};
static ExcelValue constant64 = {.type = ExcelString, .string = "Solid-fuel boiler"};
static ExcelValue constant65 = {.type = ExcelString, .string = "[2]"};
static ExcelValue constant66 = {.type = ExcelNumber, .number = 0.87};
static ExcelValue constant67 = {.type = ExcelNumber, .number = 0.13};
static ExcelValue constant68 = {.type = ExcelString, .string = "Stirling engine micro-CHP"};
static ExcelValue constant69 = {.type = ExcelString, .string = "[3]"};
static ExcelValue constant70 = {.type = ExcelNumber, .number = 0.225};
static ExcelValue constant71 = {.type = ExcelNumber, .number = 0.63};
static ExcelValue constant72 = {.type = ExcelNumber, .number = 0.145};
static ExcelValue constant73 = {.type = ExcelString, .string = "Fuel-cell micro-CHP"};
static ExcelValue constant74 = {.type = ExcelNumber, .number = 0.45};
static ExcelValue constant75 = {.type = ExcelNumber, .number = 0.1};
static ExcelValue constant76 = {.type = ExcelString, .string = "Air-source heat pump"};
static ExcelValue constant77 = {.type = ExcelString, .string = "Ground-source heat pump"};
static ExcelValue constant78 = {.type = ExcelNumber, .number = -2.0};
static ExcelValue constant79 = {.type = ExcelString, .string = "Geothermal electricity"};
static ExcelValue constant80 = {.type = ExcelNumber, .number = 0.85};
static ExcelValue constant81 = {.type = ExcelNumber, .number = 0.15};
static ExcelValue constant82 = {.type = ExcelNumber, .number = 11.0};
static ExcelValue constant83 = {.type = ExcelString, .string = "Community scale gas CHP with local district heating"};
static ExcelValue constant84 = {.type = ExcelNumber, .number = 0.38};
static ExcelValue constant85 = {.type = ExcelString, .string = "Community scale solid-fuel CHP with local district heating"};
static ExcelValue constant86 = {.type = ExcelNumber, .number = 0.17};
static ExcelValue constant87 = {.type = ExcelNumber, .number = 0.57};
static ExcelValue constant88 = {.type = ExcelNumber, .number = 0.26};
static ExcelValue constant89 = {.type = ExcelNumber, .number = 13.0};
static ExcelValue constant90 = {.type = ExcelString, .string = "Long distance district heating from large power stations"};
static ExcelValue constant91 = {.type = ExcelString, .string = "[6]"};
static ExcelValue constant92 = {.type = ExcelNumber, .number = 0.9};
static ExcelValue constant93 = {.type = ExcelNumber, .number = 137.26515207025273};
static ExcelValue constant94 = {.type = ExcelNumber, .number = 30.731004194832696};
static ExcelValue constant95 = {.type = ExcelNumber, .number = 20.487336129888465};
static ExcelValue constant96 = {.type = ExcelNumber, .number = 8.194934451955387};
static ExcelValue constant97 = {.type = ExcelString, .string = "Alpha"};
static ExcelValue constant98 = {.type = ExcelString, .string = "Beta"};
static ExcelValue constant99 = {.type = ExcelString, .string = "Gamma"};
static ExcelValue constant100 = {.type = ExcelString, .string = "ColA"};
static ExcelValue constant101 = {.type = ExcelString, .string = "ColB"};
static ExcelValue constant102 = {.type = ExcelString, .string = "Column1"};
static ExcelValue constant103 = {.type = ExcelString, .string = "A"};
static ExcelValue constant104 = {.type = ExcelString, .string = "Z"};
static ExcelValue constant105 = {.type = ExcelString, .string = "B"};
static ExcelValue constant106 = {.type = ExcelString, .string = "ZB"};
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
  return ONE;
}
static ExcelValue valuetypes_a3_variable;
ExcelValue valuetypes_a3() { if(variable_set[2] == 1) { return valuetypes_a3_variable; } else { return valuetypes_a3_default(); } }
void set_valuetypes_a3(ExcelValue newValue) { variable_set[2] = 1; valuetypes_a3_variable = newValue; }

ExcelValue valuetypes_a4_default() {
  return constant2;
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

ExcelValue formulaetypes_a1() { return constant3; }
ExcelValue formulaetypes_b1() {
  static ExcelValue result;
  if(variable_set[7] == 1) { return result;}
  result = TWO;
  variable_set[7] = 1;
  return result;
}

ExcelValue formulaetypes_a2() { return constant4; }
ExcelValue formulaetypes_b2() { return constant5; }
ExcelValue formulaetypes_a3() { return constant6; }
ExcelValue formulaetypes_b3() { return constant5; }
ExcelValue formulaetypes_a4() { return constant6; }
ExcelValue formulaetypes_b4() { return constant5; }
ExcelValue formulaetypes_a5() { return constant7; }
ExcelValue formulaetypes_b5() {
  static ExcelValue result;
  if(variable_set[15] == 1) { return result;}
  result = TWO;
  variable_set[15] = 1;
  return result;
}

ExcelValue formulaetypes_a6() { return constant8; }
ExcelValue formulaetypes_b6() { return constant9; }
ExcelValue formulaetypes_a7() { return constant10; }
ExcelValue formulaetypes_b7() { return constant9; }
ExcelValue formulaetypes_a8() { return constant10; }
ExcelValue formulaetypes_b8() { return constant9; }
ExcelValue ranges_b1() { return constant11; }
ExcelValue ranges_c1() { return constant12; }
ExcelValue ranges_a2() { return constant13; }
ExcelValue ranges_b2() {
  static ExcelValue result;
  if(variable_set[25] == 1) { return result;}
  ExcelValue array0[] = {ranges_f4(),ranges_f5(),ranges_f6()};
  result = sum(3, array0);
  variable_set[25] = 1;
  return result;
}

ExcelValue ranges_c2() {
  static ExcelValue result;
  if(variable_set[26] == 1) { return result;}
  ExcelValue array0[] = {valuetypes_a3(),valuetypes_a4()};
  result = sum(2, array0);
  variable_set[26] = 1;
  return result;
}

ExcelValue ranges_a3() { return constant14; }
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
  ExcelValue array1_ev = EXCEL_RANGE(array1,6,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[28] = 1;
  return result;
}

ExcelValue ranges_c3() {
  static ExcelValue result;
  if(variable_set[29] == 1) { return result;}
  ExcelValue array0[] = {valuetypes_a1(),valuetypes_a2(),valuetypes_a3(),valuetypes_a4(),valuetypes_a5(),valuetypes_a6()};
  result = sum(6, array0);
  variable_set[29] = 1;
  return result;
}

ExcelValue ranges_a4() { return constant15; }
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
  ExcelValue array1_ev = EXCEL_RANGE(array1,1,7);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[31] = 1;
  return result;
}

ExcelValue ranges_c4() {
  static ExcelValue result;
  if(variable_set[32] == 1) { return result;}
  result = number_or_zero(valuetypes_a4());
  variable_set[32] = 1;
  return result;
}

ExcelValue ranges_f4_default() {
  return ONE;
}
static ExcelValue ranges_f4_variable;
ExcelValue ranges_f4() { if(variable_set[33] == 1) { return ranges_f4_variable; } else { return ranges_f4_default(); } }
void set_ranges_f4(ExcelValue newValue) { variable_set[33] = 1; ranges_f4_variable = newValue; }

ExcelValue ranges_e5_default() {
  return ONE;
}
static ExcelValue ranges_e5_variable;
ExcelValue ranges_e5() { if(variable_set[34] == 1) { return ranges_e5_variable; } else { return ranges_e5_default(); } }
void set_ranges_e5(ExcelValue newValue) { variable_set[34] = 1; ranges_e5_variable = newValue; }

ExcelValue ranges_f5_default() {
  return TWO;
}
static ExcelValue ranges_f5_variable;
ExcelValue ranges_f5() { if(variable_set[35] == 1) { return ranges_f5_variable; } else { return ranges_f5_default(); } }
void set_ranges_f5(ExcelValue newValue) { variable_set[35] = 1; ranges_f5_variable = newValue; }

ExcelValue ranges_g5_default() {
  return THREE;
}
static ExcelValue ranges_g5_variable;
ExcelValue ranges_g5() { if(variable_set[36] == 1) { return ranges_g5_variable; } else { return ranges_g5_default(); } }
void set_ranges_g5(ExcelValue newValue) { variable_set[36] = 1; ranges_g5_variable = newValue; }

ExcelValue ranges_f6_default() {
  return THREE;
}
static ExcelValue ranges_f6_variable;
ExcelValue ranges_f6() { if(variable_set[37] == 1) { return ranges_f6_variable; } else { return ranges_f6_default(); } }
void set_ranges_f6(ExcelValue newValue) { variable_set[37] = 1; ranges_f6_variable = newValue; }

ExcelValue referencing_a1() { return common0(); }
ExcelValue referencing_a2() {
  static ExcelValue result;
  if(variable_set[39] == 1) { return result;}
  result = referencing_a1();
  variable_set[39] = 1;
  return result;
}

ExcelValue referencing_a4_default() {
  return TEN;
}
static ExcelValue referencing_a4_variable;
ExcelValue referencing_a4() { if(variable_set[40] == 1) { return referencing_a4_variable; } else { return referencing_a4_default(); } }
void set_referencing_a4(ExcelValue newValue) { variable_set[40] = 1; referencing_a4_variable = newValue; }

ExcelValue referencing_b4() { return common1(); }
ExcelValue referencing_c4() { return common0(); }
ExcelValue referencing_a5() {
  static ExcelValue result;
  if(variable_set[43] == 1) { return result;}
  result = THREE;
  variable_set[43] = 1;
  return result;
}

ExcelValue referencing_b8() {
  static ExcelValue result;
  if(variable_set[44] == 1) { return result;}
  result = referencing_a1();
  variable_set[44] = 1;
  return result;
}

ExcelValue referencing_b9() { return common2(); }
ExcelValue referencing_b11() { return constant16; }
ExcelValue referencing_c11() { return constant17; }
ExcelValue referencing_c15_default() {
  return ONE;
}
static ExcelValue referencing_c15_variable;
ExcelValue referencing_c15() { if(variable_set[48] == 1) { return referencing_c15_variable; } else { return referencing_c15_default(); } }
void set_referencing_c15(ExcelValue newValue) { variable_set[48] = 1; referencing_c15_variable = newValue; }

ExcelValue referencing_d15_default() {
  return TWO;
}
static ExcelValue referencing_d15_variable;
ExcelValue referencing_d15() { if(variable_set[49] == 1) { return referencing_d15_variable; } else { return referencing_d15_default(); } }
void set_referencing_d15(ExcelValue newValue) { variable_set[49] = 1; referencing_d15_variable = newValue; }

ExcelValue referencing_e15_default() {
  return THREE;
}
static ExcelValue referencing_e15_variable;
ExcelValue referencing_e15() { if(variable_set[50] == 1) { return referencing_e15_variable; } else { return referencing_e15_default(); } }
void set_referencing_e15(ExcelValue newValue) { variable_set[50] = 1; referencing_e15_variable = newValue; }

ExcelValue referencing_f15_default() {
  return FOUR;
}
static ExcelValue referencing_f15_variable;
ExcelValue referencing_f15() { if(variable_set[51] == 1) { return referencing_f15_variable; } else { return referencing_f15_default(); } }
void set_referencing_f15(ExcelValue newValue) { variable_set[51] = 1; referencing_f15_variable = newValue; }

ExcelValue referencing_c16_default() {
  return constant18;
}
static ExcelValue referencing_c16_variable;
ExcelValue referencing_c16() { if(variable_set[52] == 1) { return referencing_c16_variable; } else { return referencing_c16_default(); } }
void set_referencing_c16(ExcelValue newValue) { variable_set[52] = 1; referencing_c16_variable = newValue; }

ExcelValue referencing_d16_default() {
  return constant18;
}
static ExcelValue referencing_d16_variable;
ExcelValue referencing_d16() { if(variable_set[53] == 1) { return referencing_d16_variable; } else { return referencing_d16_default(); } }
void set_referencing_d16(ExcelValue newValue) { variable_set[53] = 1; referencing_d16_variable = newValue; }

ExcelValue referencing_e16_default() {
  return constant19;
}
static ExcelValue referencing_e16_variable;
ExcelValue referencing_e16() { if(variable_set[54] == 1) { return referencing_e16_variable; } else { return referencing_e16_default(); } }
void set_referencing_e16(ExcelValue newValue) { variable_set[54] = 1; referencing_e16_variable = newValue; }

ExcelValue referencing_f16_default() {
  return constant20;
}
static ExcelValue referencing_f16_variable;
ExcelValue referencing_f16() { if(variable_set[55] == 1) { return referencing_f16_variable; } else { return referencing_f16_default(); } }
void set_referencing_f16(ExcelValue newValue) { variable_set[55] = 1; referencing_f16_variable = newValue; }

ExcelValue referencing_c17_default() {
  return constant21;
}
static ExcelValue referencing_c17_variable;
ExcelValue referencing_c17() { if(variable_set[56] == 1) { return referencing_c17_variable; } else { return referencing_c17_default(); } }
void set_referencing_c17(ExcelValue newValue) { variable_set[56] = 1; referencing_c17_variable = newValue; }

ExcelValue referencing_d17_default() {
  return constant22;
}
static ExcelValue referencing_d17_variable;
ExcelValue referencing_d17() { if(variable_set[57] == 1) { return referencing_d17_variable; } else { return referencing_d17_default(); } }
void set_referencing_d17(ExcelValue newValue) { variable_set[57] = 1; referencing_d17_variable = newValue; }

ExcelValue referencing_e17_default() {
  return constant23;
}
static ExcelValue referencing_e17_variable;
ExcelValue referencing_e17() { if(variable_set[58] == 1) { return referencing_e17_variable; } else { return referencing_e17_default(); } }
void set_referencing_e17(ExcelValue newValue) { variable_set[58] = 1; referencing_e17_variable = newValue; }

ExcelValue referencing_f17_default() {
  return constant23;
}
static ExcelValue referencing_f17_variable;
ExcelValue referencing_f17() { if(variable_set[59] == 1) { return referencing_f17_variable; } else { return referencing_f17_default(); } }
void set_referencing_f17(ExcelValue newValue) { variable_set[59] = 1; referencing_f17_variable = newValue; }

ExcelValue referencing_c18_default() {
  return constant24;
}
static ExcelValue referencing_c18_variable;
ExcelValue referencing_c18() { if(variable_set[60] == 1) { return referencing_c18_variable; } else { return referencing_c18_default(); } }
void set_referencing_c18(ExcelValue newValue) { variable_set[60] = 1; referencing_c18_variable = newValue; }

ExcelValue referencing_d18_default() {
  return constant24;
}
static ExcelValue referencing_d18_variable;
ExcelValue referencing_d18() { if(variable_set[61] == 1) { return referencing_d18_variable; } else { return referencing_d18_default(); } }
void set_referencing_d18(ExcelValue newValue) { variable_set[61] = 1; referencing_d18_variable = newValue; }

ExcelValue referencing_e18_default() {
  return constant25;
}
static ExcelValue referencing_e18_variable;
ExcelValue referencing_e18() { if(variable_set[62] == 1) { return referencing_e18_variable; } else { return referencing_e18_default(); } }
void set_referencing_e18(ExcelValue newValue) { variable_set[62] = 1; referencing_e18_variable = newValue; }

ExcelValue referencing_f18_default() {
  return constant26;
}
static ExcelValue referencing_f18_variable;
ExcelValue referencing_f18() { if(variable_set[63] == 1) { return referencing_f18_variable; } else { return referencing_f18_default(); } }
void set_referencing_f18(ExcelValue newValue) { variable_set[63] = 1; referencing_f18_variable = newValue; }

ExcelValue referencing_c19_default() {
  return constant27;
}
static ExcelValue referencing_c19_variable;
ExcelValue referencing_c19() { if(variable_set[64] == 1) { return referencing_c19_variable; } else { return referencing_c19_default(); } }
void set_referencing_c19(ExcelValue newValue) { variable_set[64] = 1; referencing_c19_variable = newValue; }

ExcelValue referencing_d19_default() {
  return constant27;
}
static ExcelValue referencing_d19_variable;
ExcelValue referencing_d19() { if(variable_set[65] == 1) { return referencing_d19_variable; } else { return referencing_d19_default(); } }
void set_referencing_d19(ExcelValue newValue) { variable_set[65] = 1; referencing_d19_variable = newValue; }

ExcelValue referencing_e19_default() {
  return constant27;
}
static ExcelValue referencing_e19_variable;
ExcelValue referencing_e19() { if(variable_set[66] == 1) { return referencing_e19_variable; } else { return referencing_e19_default(); } }
void set_referencing_e19(ExcelValue newValue) { variable_set[66] = 1; referencing_e19_variable = newValue; }

ExcelValue referencing_f19_default() {
  return constant27;
}
static ExcelValue referencing_f19_variable;
ExcelValue referencing_f19() { if(variable_set[67] == 1) { return referencing_f19_variable; } else { return referencing_f19_default(); } }
void set_referencing_f19(ExcelValue newValue) { variable_set[67] = 1; referencing_f19_variable = newValue; }

ExcelValue referencing_c22_default() {
  return FOUR;
}
static ExcelValue referencing_c22_variable;
ExcelValue referencing_c22() { if(variable_set[68] == 1) { return referencing_c22_variable; } else { return referencing_c22_default(); } }
void set_referencing_c22(ExcelValue newValue) { variable_set[68] = 1; referencing_c22_variable = newValue; }

ExcelValue referencing_d22() {
  static ExcelValue result;
  if(variable_set[69] == 1) { return result;}
  result = excel_index(common3(),ONE,ONE);
  variable_set[69] = 1;
  return result;
}

ExcelValue referencing_d23() {
  static ExcelValue result;
  if(variable_set[70] == 1) { return result;}
  result = excel_index(common3(),TWO,ONE);
  variable_set[70] = 1;
  return result;
}

ExcelValue referencing_d24() {
  static ExcelValue result;
  if(variable_set[71] == 1) { return result;}
  result = excel_index(common3(),THREE,ONE);
  variable_set[71] = 1;
  return result;
}

ExcelValue referencing_d25() {
  static ExcelValue result;
  if(variable_set[72] == 1) { return result;}
  result = excel_index(common3(),FOUR,ONE);
  variable_set[72] = 1;
  return result;
}

ExcelValue referencing_c31() { return constant28; }
ExcelValue referencing_o31() { return constant29; }
ExcelValue referencing_f33() { return constant30; }
ExcelValue referencing_g33() { return constant31; }
ExcelValue referencing_h33() { return constant32; }
ExcelValue referencing_i33() { return constant33; }
ExcelValue referencing_j33() { return constant34; }
ExcelValue referencing_k33() { return constant35; }
ExcelValue referencing_l33() { return constant36; }
ExcelValue referencing_m33() { return constant37; }
ExcelValue referencing_n33() { return constant38; }
ExcelValue referencing_o33() { return constant39; }
ExcelValue referencing_c34() { return constant40; }
ExcelValue referencing_d34() { return constant41; }
ExcelValue referencing_e34() { return constant42; }
ExcelValue referencing_f34_default() {
  return constant43;
}
static ExcelValue referencing_f34_variable;
ExcelValue referencing_f34() { if(variable_set[88] == 1) { return referencing_f34_variable; } else { return referencing_f34_default(); } }
void set_referencing_f34(ExcelValue newValue) { variable_set[88] = 1; referencing_f34_variable = newValue; }

ExcelValue referencing_g34_default() {
  return constant44;
}
static ExcelValue referencing_g34_variable;
ExcelValue referencing_g34() { if(variable_set[89] == 1) { return referencing_g34_variable; } else { return referencing_g34_default(); } }
void set_referencing_g34(ExcelValue newValue) { variable_set[89] = 1; referencing_g34_variable = newValue; }

ExcelValue referencing_h34_default() {
  return constant45;
}
static ExcelValue referencing_h34_variable;
ExcelValue referencing_h34() { if(variable_set[90] == 1) { return referencing_h34_variable; } else { return referencing_h34_default(); } }
void set_referencing_h34(ExcelValue newValue) { variable_set[90] = 1; referencing_h34_variable = newValue; }

ExcelValue referencing_i34_default() {
  return constant46;
}
static ExcelValue referencing_i34_variable;
ExcelValue referencing_i34() { if(variable_set[91] == 1) { return referencing_i34_variable; } else { return referencing_i34_default(); } }
void set_referencing_i34(ExcelValue newValue) { variable_set[91] = 1; referencing_i34_variable = newValue; }

ExcelValue referencing_j34_default() {
  return constant47;
}
static ExcelValue referencing_j34_variable;
ExcelValue referencing_j34() { if(variable_set[92] == 1) { return referencing_j34_variable; } else { return referencing_j34_default(); } }
void set_referencing_j34(ExcelValue newValue) { variable_set[92] = 1; referencing_j34_variable = newValue; }

ExcelValue referencing_k34_default() {
  return constant48;
}
static ExcelValue referencing_k34_variable;
ExcelValue referencing_k34() { if(variable_set[93] == 1) { return referencing_k34_variable; } else { return referencing_k34_default(); } }
void set_referencing_k34(ExcelValue newValue) { variable_set[93] = 1; referencing_k34_variable = newValue; }

ExcelValue referencing_l34_default() {
  return constant49;
}
static ExcelValue referencing_l34_variable;
ExcelValue referencing_l34() { if(variable_set[94] == 1) { return referencing_l34_variable; } else { return referencing_l34_default(); } }
void set_referencing_l34(ExcelValue newValue) { variable_set[94] = 1; referencing_l34_variable = newValue; }

ExcelValue referencing_m34_default() {
  return constant50;
}
static ExcelValue referencing_m34_variable;
ExcelValue referencing_m34() { if(variable_set[95] == 1) { return referencing_m34_variable; } else { return referencing_m34_default(); } }
void set_referencing_m34(ExcelValue newValue) { variable_set[95] = 1; referencing_m34_variable = newValue; }

ExcelValue referencing_n34_default() {
  return constant51;
}
static ExcelValue referencing_n34_variable;
ExcelValue referencing_n34() { if(variable_set[96] == 1) { return referencing_n34_variable; } else { return referencing_n34_default(); } }
void set_referencing_n34(ExcelValue newValue) { variable_set[96] = 1; referencing_n34_variable = newValue; }

ExcelValue referencing_c35() {
  static ExcelValue result;
  if(variable_set[97] == 1) { return result;}
  result = ONE;
  variable_set[97] = 1;
  return result;
}

ExcelValue referencing_d35() { return constant52; }
ExcelValue referencing_j35_default() {
  return constant53;
}
static ExcelValue referencing_j35_variable;
ExcelValue referencing_j35() { if(variable_set[99] == 1) { return referencing_j35_variable; } else { return referencing_j35_default(); } }
void set_referencing_j35(ExcelValue newValue) { variable_set[99] = 1; referencing_j35_variable = newValue; }

ExcelValue referencing_m35_default() {
  return constant54;
}
static ExcelValue referencing_m35_variable;
ExcelValue referencing_m35() { if(variable_set[100] == 1) { return referencing_m35_variable; } else { return referencing_m35_default(); } }
void set_referencing_m35(ExcelValue newValue) { variable_set[100] = 1; referencing_m35_variable = newValue; }

ExcelValue referencing_n35_default() {
  return constant55;
}
static ExcelValue referencing_n35_variable;
ExcelValue referencing_n35() { if(variable_set[101] == 1) { return referencing_n35_variable; } else { return referencing_n35_default(); } }
void set_referencing_n35(ExcelValue newValue) { variable_set[101] = 1; referencing_n35_variable = newValue; }

ExcelValue referencing_o35() {
  static ExcelValue result;
  if(variable_set[102] == 1) { return result;}
  result = ZERO;
  variable_set[102] = 1;
  return result;
}

ExcelValue referencing_c36() {
  static ExcelValue result;
  if(variable_set[103] == 1) { return result;}
  result = TWO;
  variable_set[103] = 1;
  return result;
}

ExcelValue referencing_d36() { return constant56; }
ExcelValue referencing_j36_default() {
  return constant53;
}
static ExcelValue referencing_j36_variable;
ExcelValue referencing_j36() { if(variable_set[105] == 1) { return referencing_j36_variable; } else { return referencing_j36_default(); } }
void set_referencing_j36(ExcelValue newValue) { variable_set[105] = 1; referencing_j36_variable = newValue; }

ExcelValue referencing_m36_default() {
  return constant57;
}
static ExcelValue referencing_m36_variable;
ExcelValue referencing_m36() { if(variable_set[106] == 1) { return referencing_m36_variable; } else { return referencing_m36_default(); } }
void set_referencing_m36(ExcelValue newValue) { variable_set[106] = 1; referencing_m36_variable = newValue; }

ExcelValue referencing_n36_default() {
  return constant58;
}
static ExcelValue referencing_n36_variable;
ExcelValue referencing_n36() { if(variable_set[107] == 1) { return referencing_n36_variable; } else { return referencing_n36_default(); } }
void set_referencing_n36(ExcelValue newValue) { variable_set[107] = 1; referencing_n36_variable = newValue; }

ExcelValue referencing_o36() {
  static ExcelValue result;
  if(variable_set[108] == 1) { return result;}
  result = ZERO;
  variable_set[108] = 1;
  return result;
}

ExcelValue referencing_c37() {
  static ExcelValue result;
  if(variable_set[109] == 1) { return result;}
  result = THREE;
  variable_set[109] = 1;
  return result;
}

ExcelValue referencing_d37() { return constant59; }
ExcelValue referencing_f37_default() {
  return constant53;
}
static ExcelValue referencing_f37_variable;
ExcelValue referencing_f37() { if(variable_set[111] == 1) { return referencing_f37_variable; } else { return referencing_f37_default(); } }
void set_referencing_f37(ExcelValue newValue) { variable_set[111] = 1; referencing_f37_variable = newValue; }

ExcelValue referencing_m37_default() {
  return ONE;
}
static ExcelValue referencing_m37_variable;
ExcelValue referencing_m37() { if(variable_set[112] == 1) { return referencing_m37_variable; } else { return referencing_m37_default(); } }
void set_referencing_m37(ExcelValue newValue) { variable_set[112] = 1; referencing_m37_variable = newValue; }

ExcelValue referencing_n37_default() {
  return ZERO;
}
static ExcelValue referencing_n37_variable;
ExcelValue referencing_n37() { if(variable_set[113] == 1) { return referencing_n37_variable; } else { return referencing_n37_default(); } }
void set_referencing_n37(ExcelValue newValue) { variable_set[113] = 1; referencing_n37_variable = newValue; }

ExcelValue referencing_o37() {
  static ExcelValue result;
  if(variable_set[114] == 1) { return result;}
  result = ZERO;
  variable_set[114] = 1;
  return result;
}

ExcelValue referencing_c38() {
  static ExcelValue result;
  if(variable_set[115] == 1) { return result;}
  result = FOUR;
  variable_set[115] = 1;
  return result;
}

ExcelValue referencing_d38() { return constant60; }
ExcelValue referencing_i38_default() {
  return constant53;
}
static ExcelValue referencing_i38_variable;
ExcelValue referencing_i38() { if(variable_set[117] == 1) { return referencing_i38_variable; } else { return referencing_i38_default(); } }
void set_referencing_i38(ExcelValue newValue) { variable_set[117] = 1; referencing_i38_variable = newValue; }

ExcelValue referencing_m38_default() {
  return constant61;
}
static ExcelValue referencing_m38_variable;
ExcelValue referencing_m38() { if(variable_set[118] == 1) { return referencing_m38_variable; } else { return referencing_m38_default(); } }
void set_referencing_m38(ExcelValue newValue) { variable_set[118] = 1; referencing_m38_variable = newValue; }

ExcelValue referencing_n38_default() {
  return constant62;
}
static ExcelValue referencing_n38_variable;
ExcelValue referencing_n38() { if(variable_set[119] == 1) { return referencing_n38_variable; } else { return referencing_n38_default(); } }
void set_referencing_n38(ExcelValue newValue) { variable_set[119] = 1; referencing_n38_variable = newValue; }

ExcelValue referencing_o38() { return constant63; }
ExcelValue referencing_c39() {
  static ExcelValue result;
  if(variable_set[121] == 1) { return result;}
  result = FIVE;
  variable_set[121] = 1;
  return result;
}

ExcelValue referencing_d39() { return constant64; }
ExcelValue referencing_e39() { return constant65; }
ExcelValue referencing_h39_default() {
  return constant53;
}
static ExcelValue referencing_h39_variable;
ExcelValue referencing_h39() { if(variable_set[124] == 1) { return referencing_h39_variable; } else { return referencing_h39_default(); } }
void set_referencing_h39(ExcelValue newValue) { variable_set[124] = 1; referencing_h39_variable = newValue; }

ExcelValue referencing_m39_default() {
  return constant66;
}
static ExcelValue referencing_m39_variable;
ExcelValue referencing_m39() { if(variable_set[125] == 1) { return referencing_m39_variable; } else { return referencing_m39_default(); } }
void set_referencing_m39(ExcelValue newValue) { variable_set[125] = 1; referencing_m39_variable = newValue; }

ExcelValue referencing_n39_default() {
  return constant67;
}
static ExcelValue referencing_n39_variable;
ExcelValue referencing_n39() { if(variable_set[126] == 1) { return referencing_n39_variable; } else { return referencing_n39_default(); } }
void set_referencing_n39(ExcelValue newValue) { variable_set[126] = 1; referencing_n39_variable = newValue; }

ExcelValue referencing_o39() {
  static ExcelValue result;
  if(variable_set[127] == 1) { return result;}
  result = ZERO;
  variable_set[127] = 1;
  return result;
}

ExcelValue referencing_c40() {
  static ExcelValue result;
  if(variable_set[128] == 1) { return result;}
  result = SIX;
  variable_set[128] = 1;
  return result;
}

ExcelValue referencing_d40() { return constant68; }
ExcelValue referencing_e40() { return constant69; }
ExcelValue referencing_g40_default() {
  return constant70;
}
static ExcelValue referencing_g40_variable;
ExcelValue referencing_g40() { if(variable_set[131] == 1) { return referencing_g40_variable; } else { return referencing_g40_default(); } }
void set_referencing_g40(ExcelValue newValue) { variable_set[131] = 1; referencing_g40_variable = newValue; }

ExcelValue referencing_j40_default() {
  return constant53;
}
static ExcelValue referencing_j40_variable;
ExcelValue referencing_j40() { if(variable_set[132] == 1) { return referencing_j40_variable; } else { return referencing_j40_default(); } }
void set_referencing_j40(ExcelValue newValue) { variable_set[132] = 1; referencing_j40_variable = newValue; }

ExcelValue referencing_m40_default() {
  return constant71;
}
static ExcelValue referencing_m40_variable;
ExcelValue referencing_m40() { if(variable_set[133] == 1) { return referencing_m40_variable; } else { return referencing_m40_default(); } }
void set_referencing_m40(ExcelValue newValue) { variable_set[133] = 1; referencing_m40_variable = newValue; }

ExcelValue referencing_n40_default() {
  return constant72;
}
static ExcelValue referencing_n40_variable;
ExcelValue referencing_n40() { if(variable_set[134] == 1) { return referencing_n40_variable; } else { return referencing_n40_default(); } }
void set_referencing_n40(ExcelValue newValue) { variable_set[134] = 1; referencing_n40_variable = newValue; }

ExcelValue referencing_o40() {
  static ExcelValue result;
  if(variable_set[135] == 1) { return result;}
  result = ZERO;
  variable_set[135] = 1;
  return result;
}

ExcelValue referencing_c41() {
  static ExcelValue result;
  if(variable_set[136] == 1) { return result;}
  result = SEVEN;
  variable_set[136] = 1;
  return result;
}

ExcelValue referencing_d41() { return constant73; }
ExcelValue referencing_e41() { return constant69; }
ExcelValue referencing_g41_default() {
  return constant74;
}
static ExcelValue referencing_g41_variable;
ExcelValue referencing_g41() { if(variable_set[139] == 1) { return referencing_g41_variable; } else { return referencing_g41_default(); } }
void set_referencing_g41(ExcelValue newValue) { variable_set[139] = 1; referencing_g41_variable = newValue; }

ExcelValue referencing_j41_default() {
  return constant53;
}
static ExcelValue referencing_j41_variable;
ExcelValue referencing_j41() { if(variable_set[140] == 1) { return referencing_j41_variable; } else { return referencing_j41_default(); } }
void set_referencing_j41(ExcelValue newValue) { variable_set[140] = 1; referencing_j41_variable = newValue; }

ExcelValue referencing_m41_default() {
  return constant74;
}
static ExcelValue referencing_m41_variable;
ExcelValue referencing_m41() { if(variable_set[141] == 1) { return referencing_m41_variable; } else { return referencing_m41_default(); } }
void set_referencing_m41(ExcelValue newValue) { variable_set[141] = 1; referencing_m41_variable = newValue; }

ExcelValue referencing_n41_default() {
  return constant75;
}
static ExcelValue referencing_n41_variable;
ExcelValue referencing_n41() { if(variable_set[142] == 1) { return referencing_n41_variable; } else { return referencing_n41_default(); } }
void set_referencing_n41(ExcelValue newValue) { variable_set[142] = 1; referencing_n41_variable = newValue; }

ExcelValue referencing_o41() {
  static ExcelValue result;
  if(variable_set[143] == 1) { return result;}
  result = ZERO;
  variable_set[143] = 1;
  return result;
}

ExcelValue referencing_c42() {
  static ExcelValue result;
  if(variable_set[144] == 1) { return result;}
  result = EIGHT;
  variable_set[144] = 1;
  return result;
}

ExcelValue referencing_d42() { return constant76; }
ExcelValue referencing_f42_default() {
  return constant53;
}
static ExcelValue referencing_f42_variable;
ExcelValue referencing_f42() { if(variable_set[146] == 1) { return referencing_f42_variable; } else { return referencing_f42_default(); } }
void set_referencing_f42(ExcelValue newValue) { variable_set[146] = 1; referencing_f42_variable = newValue; }

ExcelValue referencing_l42_default() {
  return constant53;
}
static ExcelValue referencing_l42_variable;
ExcelValue referencing_l42() { if(variable_set[147] == 1) { return referencing_l42_variable; } else { return referencing_l42_default(); } }
void set_referencing_l42(ExcelValue newValue) { variable_set[147] = 1; referencing_l42_variable = newValue; }

ExcelValue referencing_m42_default() {
  return TWO;
}
static ExcelValue referencing_m42_variable;
ExcelValue referencing_m42() { if(variable_set[148] == 1) { return referencing_m42_variable; } else { return referencing_m42_default(); } }
void set_referencing_m42(ExcelValue newValue) { variable_set[148] = 1; referencing_m42_variable = newValue; }

ExcelValue referencing_o42() {
  static ExcelValue result;
  if(variable_set[149] == 1) { return result;}
  result = ZERO;
  variable_set[149] = 1;
  return result;
}

ExcelValue referencing_c43() {
  static ExcelValue result;
  if(variable_set[150] == 1) { return result;}
  result = NINE;
  variable_set[150] = 1;
  return result;
}

ExcelValue referencing_d43() { return constant77; }
ExcelValue referencing_f43_default() {
  return constant53;
}
static ExcelValue referencing_f43_variable;
ExcelValue referencing_f43() { if(variable_set[152] == 1) { return referencing_f43_variable; } else { return referencing_f43_default(); } }
void set_referencing_f43(ExcelValue newValue) { variable_set[152] = 1; referencing_f43_variable = newValue; }

ExcelValue referencing_l43_default() {
  return constant78;
}
static ExcelValue referencing_l43_variable;
ExcelValue referencing_l43() { if(variable_set[153] == 1) { return referencing_l43_variable; } else { return referencing_l43_default(); } }
void set_referencing_l43(ExcelValue newValue) { variable_set[153] = 1; referencing_l43_variable = newValue; }

ExcelValue referencing_m43_default() {
  return THREE;
}
static ExcelValue referencing_m43_variable;
ExcelValue referencing_m43() { if(variable_set[154] == 1) { return referencing_m43_variable; } else { return referencing_m43_default(); } }
void set_referencing_m43(ExcelValue newValue) { variable_set[154] = 1; referencing_m43_variable = newValue; }

ExcelValue referencing_o43() {
  static ExcelValue result;
  if(variable_set[155] == 1) { return result;}
  result = ZERO;
  variable_set[155] = 1;
  return result;
}

ExcelValue referencing_c44() {
  static ExcelValue result;
  if(variable_set[156] == 1) { return result;}
  result = TEN;
  variable_set[156] = 1;
  return result;
}

ExcelValue referencing_d44() { return constant79; }
ExcelValue referencing_l44_default() {
  return constant53;
}
static ExcelValue referencing_l44_variable;
ExcelValue referencing_l44() { if(variable_set[158] == 1) { return referencing_l44_variable; } else { return referencing_l44_default(); } }
void set_referencing_l44(ExcelValue newValue) { variable_set[158] = 1; referencing_l44_variable = newValue; }

ExcelValue referencing_m44_default() {
  return constant80;
}
static ExcelValue referencing_m44_variable;
ExcelValue referencing_m44() { if(variable_set[159] == 1) { return referencing_m44_variable; } else { return referencing_m44_default(); } }
void set_referencing_m44(ExcelValue newValue) { variable_set[159] = 1; referencing_m44_variable = newValue; }

ExcelValue referencing_n44_default() {
  return constant81;
}
static ExcelValue referencing_n44_variable;
ExcelValue referencing_n44() { if(variable_set[160] == 1) { return referencing_n44_variable; } else { return referencing_n44_default(); } }
void set_referencing_n44(ExcelValue newValue) { variable_set[160] = 1; referencing_n44_variable = newValue; }

ExcelValue referencing_o44() {
  static ExcelValue result;
  if(variable_set[161] == 1) { return result;}
  result = ZERO;
  variable_set[161] = 1;
  return result;
}

ExcelValue referencing_c45() { return constant82; }
ExcelValue referencing_d45() { return constant83; }
ExcelValue referencing_g45_default() {
  return constant84;
}
static ExcelValue referencing_g45_variable;
ExcelValue referencing_g45() { if(variable_set[164] == 1) { return referencing_g45_variable; } else { return referencing_g45_default(); } }
void set_referencing_g45(ExcelValue newValue) { variable_set[164] = 1; referencing_g45_variable = newValue; }

ExcelValue referencing_j45_default() {
  return constant53;
}
static ExcelValue referencing_j45_variable;
ExcelValue referencing_j45() { if(variable_set[165] == 1) { return referencing_j45_variable; } else { return referencing_j45_default(); } }
void set_referencing_j45(ExcelValue newValue) { variable_set[165] = 1; referencing_j45_variable = newValue; }

ExcelValue referencing_m45_default() {
  return constant84;
}
static ExcelValue referencing_m45_variable;
ExcelValue referencing_m45() { if(variable_set[166] == 1) { return referencing_m45_variable; } else { return referencing_m45_default(); } }
void set_referencing_m45(ExcelValue newValue) { variable_set[166] = 1; referencing_m45_variable = newValue; }

ExcelValue referencing_n45_default() {
  return constant55;
}
static ExcelValue referencing_n45_variable;
ExcelValue referencing_n45() { if(variable_set[167] == 1) { return referencing_n45_variable; } else { return referencing_n45_default(); } }
void set_referencing_n45(ExcelValue newValue) { variable_set[167] = 1; referencing_n45_variable = newValue; }

ExcelValue referencing_o45() {
  static ExcelValue result;
  if(variable_set[168] == 1) { return result;}
  result = ZERO;
  variable_set[168] = 1;
  return result;
}

ExcelValue referencing_c46() { return constant22; }
ExcelValue referencing_d46() { return constant85; }
ExcelValue referencing_g46_default() {
  return constant86;
}
static ExcelValue referencing_g46_variable;
ExcelValue referencing_g46() { if(variable_set[171] == 1) { return referencing_g46_variable; } else { return referencing_g46_default(); } }
void set_referencing_g46(ExcelValue newValue) { variable_set[171] = 1; referencing_g46_variable = newValue; }

ExcelValue referencing_h46_default() {
  return constant53;
}
static ExcelValue referencing_h46_variable;
ExcelValue referencing_h46() { if(variable_set[172] == 1) { return referencing_h46_variable; } else { return referencing_h46_default(); } }
void set_referencing_h46(ExcelValue newValue) { variable_set[172] = 1; referencing_h46_variable = newValue; }

ExcelValue referencing_m46_default() {
  return constant87;
}
static ExcelValue referencing_m46_variable;
ExcelValue referencing_m46() { if(variable_set[173] == 1) { return referencing_m46_variable; } else { return referencing_m46_default(); } }
void set_referencing_m46(ExcelValue newValue) { variable_set[173] = 1; referencing_m46_variable = newValue; }

ExcelValue referencing_n46_default() {
  return constant88;
}
static ExcelValue referencing_n46_variable;
ExcelValue referencing_n46() { if(variable_set[174] == 1) { return referencing_n46_variable; } else { return referencing_n46_default(); } }
void set_referencing_n46(ExcelValue newValue) { variable_set[174] = 1; referencing_n46_variable = newValue; }

ExcelValue referencing_o46() {
  static ExcelValue result;
  if(variable_set[175] == 1) { return result;}
  result = ZERO;
  variable_set[175] = 1;
  return result;
}

ExcelValue referencing_c47() { return constant89; }
ExcelValue referencing_d47() { return constant90; }
ExcelValue referencing_e47() { return constant91; }
ExcelValue referencing_k47_default() {
  return constant53;
}
static ExcelValue referencing_k47_variable;
ExcelValue referencing_k47() { if(variable_set[179] == 1) { return referencing_k47_variable; } else { return referencing_k47_default(); } }
void set_referencing_k47(ExcelValue newValue) { variable_set[179] = 1; referencing_k47_variable = newValue; }

ExcelValue referencing_m47_default() {
  return constant92;
}
static ExcelValue referencing_m47_variable;
ExcelValue referencing_m47() { if(variable_set[180] == 1) { return referencing_m47_variable; } else { return referencing_m47_default(); } }
void set_referencing_m47(ExcelValue newValue) { variable_set[180] = 1; referencing_m47_variable = newValue; }

ExcelValue referencing_n47_default() {
  return constant75;
}
static ExcelValue referencing_n47_variable;
ExcelValue referencing_n47() { if(variable_set[181] == 1) { return referencing_n47_variable; } else { return referencing_n47_default(); } }
void set_referencing_n47(ExcelValue newValue) { variable_set[181] = 1; referencing_n47_variable = newValue; }

ExcelValue referencing_o47() {
  static ExcelValue result;
  if(variable_set[182] == 1) { return result;}
  result = ZERO;
  variable_set[182] = 1;
  return result;
}

ExcelValue referencing_d50() { return constant52; }
ExcelValue referencing_g50_default() {
  return constant93;
}
static ExcelValue referencing_g50_variable;
ExcelValue referencing_g50() { if(variable_set[184] == 1) { return referencing_g50_variable; } else { return referencing_g50_default(); } }
void set_referencing_g50(ExcelValue newValue) { variable_set[184] = 1; referencing_g50_variable = newValue; }

ExcelValue referencing_d51() { return constant56; }
ExcelValue referencing_g51_default() {
  return constant94;
}
static ExcelValue referencing_g51_variable;
ExcelValue referencing_g51() { if(variable_set[186] == 1) { return referencing_g51_variable; } else { return referencing_g51_default(); } }
void set_referencing_g51(ExcelValue newValue) { variable_set[186] = 1; referencing_g51_variable = newValue; }

ExcelValue referencing_d52() { return constant59; }
ExcelValue referencing_g52_default() {
  return constant95;
}
static ExcelValue referencing_g52_variable;
ExcelValue referencing_g52() { if(variable_set[188] == 1) { return referencing_g52_variable; } else { return referencing_g52_default(); } }
void set_referencing_g52(ExcelValue newValue) { variable_set[188] = 1; referencing_g52_variable = newValue; }

ExcelValue referencing_d53() { return constant60; }
ExcelValue referencing_g53_default() {
  return constant96;
}
static ExcelValue referencing_g53_variable;
ExcelValue referencing_g53() { if(variable_set[190] == 1) { return referencing_g53_variable; } else { return referencing_g53_default(); } }
void set_referencing_g53(ExcelValue newValue) { variable_set[190] = 1; referencing_g53_variable = newValue; }

ExcelValue referencing_d54() { return constant64; }
ExcelValue referencing_g54_default() {
  return constant96;
}
static ExcelValue referencing_g54_variable;
ExcelValue referencing_g54() { if(variable_set[192] == 1) { return referencing_g54_variable; } else { return referencing_g54_default(); } }
void set_referencing_g54(ExcelValue newValue) { variable_set[192] = 1; referencing_g54_variable = newValue; }

ExcelValue referencing_d55() { return constant68; }
ExcelValue referencing_g55_default() {
  return ZERO;
}
static ExcelValue referencing_g55_variable;
ExcelValue referencing_g55() { if(variable_set[194] == 1) { return referencing_g55_variable; } else { return referencing_g55_default(); } }
void set_referencing_g55(ExcelValue newValue) { variable_set[194] = 1; referencing_g55_variable = newValue; }

ExcelValue referencing_d56() { return constant73; }
ExcelValue referencing_g56_default() {
  return ZERO;
}
static ExcelValue referencing_g56_variable;
ExcelValue referencing_g56() { if(variable_set[196] == 1) { return referencing_g56_variable; } else { return referencing_g56_default(); } }
void set_referencing_g56(ExcelValue newValue) { variable_set[196] = 1; referencing_g56_variable = newValue; }

ExcelValue referencing_d57() { return constant76; }
ExcelValue referencing_g57_default() {
  return ZERO;
}
static ExcelValue referencing_g57_variable;
ExcelValue referencing_g57() { if(variable_set[198] == 1) { return referencing_g57_variable; } else { return referencing_g57_default(); } }
void set_referencing_g57(ExcelValue newValue) { variable_set[198] = 1; referencing_g57_variable = newValue; }

ExcelValue referencing_d58() { return constant77; }
ExcelValue referencing_g58_default() {
  return ZERO;
}
static ExcelValue referencing_g58_variable;
ExcelValue referencing_g58() { if(variable_set[200] == 1) { return referencing_g58_variable; } else { return referencing_g58_default(); } }
void set_referencing_g58(ExcelValue newValue) { variable_set[200] = 1; referencing_g58_variable = newValue; }

ExcelValue referencing_d59() { return constant79; }
ExcelValue referencing_g59_default() {
  return ZERO;
}
static ExcelValue referencing_g59_variable;
ExcelValue referencing_g59() { if(variable_set[202] == 1) { return referencing_g59_variable; } else { return referencing_g59_default(); } }
void set_referencing_g59(ExcelValue newValue) { variable_set[202] = 1; referencing_g59_variable = newValue; }

ExcelValue referencing_d60() { return constant83; }
ExcelValue referencing_g60_default() {
  return ZERO;
}
static ExcelValue referencing_g60_variable;
ExcelValue referencing_g60() { if(variable_set[204] == 1) { return referencing_g60_variable; } else { return referencing_g60_default(); } }
void set_referencing_g60(ExcelValue newValue) { variable_set[204] = 1; referencing_g60_variable = newValue; }

ExcelValue referencing_d61() { return constant85; }
ExcelValue referencing_g61_default() {
  return ZERO;
}
static ExcelValue referencing_g61_variable;
ExcelValue referencing_g61() { if(variable_set[206] == 1) { return referencing_g61_variable; } else { return referencing_g61_default(); } }
void set_referencing_g61(ExcelValue newValue) { variable_set[206] = 1; referencing_g61_variable = newValue; }

ExcelValue referencing_d62() { return constant90; }
ExcelValue referencing_g62_default() {
  return ZERO;
}
static ExcelValue referencing_g62_variable;
ExcelValue referencing_g62() { if(variable_set[208] == 1) { return referencing_g62_variable; } else { return referencing_g62_default(); } }
void set_referencing_g62(ExcelValue newValue) { variable_set[208] = 1; referencing_g62_variable = newValue; }

ExcelValue referencing_d64_default() {
  return constant50;
}
static ExcelValue referencing_d64_variable;
ExcelValue referencing_d64() { if(variable_set[209] == 1) { return referencing_d64_variable; } else { return referencing_d64_default(); } }
void set_referencing_d64(ExcelValue newValue) { variable_set[209] = 1; referencing_d64_variable = newValue; }

ExcelValue referencing_e64() { return constant37; }
ExcelValue referencing_h64() {
  static ExcelValue result;
  if(variable_set[211] == 1) { return result;}
  static ExcelValue array1[117];
  array1[0] = ZERO;
  array1[1] = ZERO;
  array1[2] = ZERO;
  array1[3] = ZERO;
  array1[4] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_j34())),referencing_j35());
  array1[5] = ZERO;
  array1[6] = ZERO;
  array1[7] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_m34())),referencing_m35());
  array1[8] = multiply(multiply(divide(referencing_g50(),referencing_m35()),excel_equal(referencing_d64(),referencing_n34())),referencing_n35());
  array1[9] = ZERO;
  array1[10] = ZERO;
  array1[11] = ZERO;
  array1[12] = ZERO;
  array1[13] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_j34())),referencing_j36());
  array1[14] = ZERO;
  array1[15] = ZERO;
  array1[16] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_m34())),referencing_m36());
  array1[17] = multiply(multiply(divide(referencing_g51(),referencing_m36()),excel_equal(referencing_d64(),referencing_n34())),referencing_n36());
  array1[18] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_f34())),referencing_f37());
  array1[19] = ZERO;
  array1[20] = ZERO;
  array1[21] = ZERO;
  array1[22] = ZERO;
  array1[23] = ZERO;
  array1[24] = ZERO;
  array1[25] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_m34())),referencing_m37());
  array1[26] = multiply(multiply(divide(referencing_g52(),referencing_m37()),excel_equal(referencing_d64(),referencing_n34())),referencing_n37());
  array1[27] = ZERO;
  array1[28] = ZERO;
  array1[29] = ZERO;
  array1[30] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_i34())),referencing_i38());
  array1[31] = ZERO;
  array1[32] = ZERO;
  array1[33] = ZERO;
  array1[34] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_m34())),referencing_m38());
  array1[35] = multiply(multiply(divide(referencing_g53(),referencing_m38()),excel_equal(referencing_d64(),referencing_n34())),referencing_n38());
  array1[36] = ZERO;
  array1[37] = ZERO;
  array1[38] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_h34())),referencing_h39());
  array1[39] = ZERO;
  array1[40] = ZERO;
  array1[41] = ZERO;
  array1[42] = ZERO;
  array1[43] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_m34())),referencing_m39());
  array1[44] = multiply(multiply(divide(referencing_g54(),referencing_m39()),excel_equal(referencing_d64(),referencing_n34())),referencing_n39());
  array1[45] = ZERO;
  array1[46] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_g34())),referencing_g40());
  array1[47] = ZERO;
  array1[48] = ZERO;
  array1[49] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_j34())),referencing_j40());
  array1[50] = ZERO;
  array1[51] = ZERO;
  array1[52] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_m34())),referencing_m40());
  array1[53] = multiply(multiply(divide(referencing_g55(),referencing_m40()),excel_equal(referencing_d64(),referencing_n34())),referencing_n40());
  array1[54] = ZERO;
  array1[55] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_g34())),referencing_g41());
  array1[56] = ZERO;
  array1[57] = ZERO;
  array1[58] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_j34())),referencing_j41());
  array1[59] = ZERO;
  array1[60] = ZERO;
  array1[61] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_m34())),referencing_m41());
  array1[62] = multiply(multiply(divide(referencing_g56(),referencing_m41()),excel_equal(referencing_d64(),referencing_n34())),referencing_n41());
  array1[63] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_f34())),referencing_f42());
  array1[64] = ZERO;
  array1[65] = ZERO;
  array1[66] = ZERO;
  array1[67] = ZERO;
  array1[68] = ZERO;
  array1[69] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_l34())),referencing_l42());
  array1[70] = multiply(multiply(divide(referencing_g57(),referencing_m42()),excel_equal(referencing_d64(),referencing_m34())),referencing_m42());
  array1[71] = ZERO;
  array1[72] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_f34())),referencing_f43());
  array1[73] = ZERO;
  array1[74] = ZERO;
  array1[75] = ZERO;
  array1[76] = ZERO;
  array1[77] = ZERO;
  array1[78] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_l34())),referencing_l43());
  array1[79] = multiply(multiply(divide(referencing_g58(),referencing_m43()),excel_equal(referencing_d64(),referencing_m34())),referencing_m43());
  array1[80] = ZERO;
  array1[81] = ZERO;
  array1[82] = ZERO;
  array1[83] = ZERO;
  array1[84] = ZERO;
  array1[85] = ZERO;
  array1[86] = ZERO;
  array1[87] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_l34())),referencing_l44());
  array1[88] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_m34())),referencing_m44());
  array1[89] = multiply(multiply(divide(referencing_g59(),referencing_m44()),excel_equal(referencing_d64(),referencing_n34())),referencing_n44());
  array1[90] = ZERO;
  array1[91] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_g34())),referencing_g45());
  array1[92] = ZERO;
  array1[93] = ZERO;
  array1[94] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_j34())),referencing_j45());
  array1[95] = ZERO;
  array1[96] = ZERO;
  array1[97] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_m34())),referencing_m45());
  array1[98] = multiply(multiply(divide(referencing_g60(),referencing_m45()),excel_equal(referencing_d64(),referencing_n34())),referencing_n45());
  array1[99] = ZERO;
  array1[100] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_g34())),referencing_g46());
  array1[101] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_h34())),referencing_h46());
  array1[102] = ZERO;
  array1[103] = ZERO;
  array1[104] = ZERO;
  array1[105] = ZERO;
  array1[106] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_m34())),referencing_m46());
  array1[107] = multiply(multiply(divide(referencing_g61(),referencing_m46()),excel_equal(referencing_d64(),referencing_n34())),referencing_n46());
  array1[108] = ZERO;
  array1[109] = ZERO;
  array1[110] = ZERO;
  array1[111] = ZERO;
  array1[112] = ZERO;
  array1[113] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_k34())),referencing_k47());
  array1[114] = ZERO;
  array1[115] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_m34())),referencing_m47());
  array1[116] = multiply(multiply(divide(referencing_g62(),referencing_m47()),excel_equal(referencing_d64(),referencing_n34())),referencing_n47());
  ExcelValue array1_ev = EXCEL_RANGE(array1,13,9);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[211] = 1;
  return result;
}

ExcelValue referencing_e68_default() {
  return constant97;
}
static ExcelValue referencing_e68_variable;
ExcelValue referencing_e68() { if(variable_set[212] == 1) { return referencing_e68_variable; } else { return referencing_e68_default(); } }
void set_referencing_e68(ExcelValue newValue) { variable_set[212] = 1; referencing_e68_variable = newValue; }

ExcelValue referencing_f68_default() {
  return ONE;
}
static ExcelValue referencing_f68_variable;
ExcelValue referencing_f68() { if(variable_set[213] == 1) { return referencing_f68_variable; } else { return referencing_f68_default(); } }
void set_referencing_f68(ExcelValue newValue) { variable_set[213] = 1; referencing_f68_variable = newValue; }

ExcelValue referencing_e69_default() {
  return constant98;
}
static ExcelValue referencing_e69_variable;
ExcelValue referencing_e69() { if(variable_set[214] == 1) { return referencing_e69_variable; } else { return referencing_e69_default(); } }
void set_referencing_e69(ExcelValue newValue) { variable_set[214] = 1; referencing_e69_variable = newValue; }

ExcelValue referencing_f69_default() {
  return TWO;
}
static ExcelValue referencing_f69_variable;
ExcelValue referencing_f69() { if(variable_set[215] == 1) { return referencing_f69_variable; } else { return referencing_f69_default(); } }
void set_referencing_f69(ExcelValue newValue) { variable_set[215] = 1; referencing_f69_variable = newValue; }

ExcelValue referencing_e70_default() {
  return constant99;
}
static ExcelValue referencing_e70_variable;
ExcelValue referencing_e70() { if(variable_set[216] == 1) { return referencing_e70_variable; } else { return referencing_e70_default(); } }
void set_referencing_e70(ExcelValue newValue) { variable_set[216] = 1; referencing_e70_variable = newValue; }

ExcelValue referencing_f70_default() {
  return THREE;
}
static ExcelValue referencing_f70_variable;
ExcelValue referencing_f70() { if(variable_set[217] == 1) { return referencing_f70_variable; } else { return referencing_f70_default(); } }
void set_referencing_f70(ExcelValue newValue) { variable_set[217] = 1; referencing_f70_variable = newValue; }

ExcelValue referencing_e72_default() {
  return constant98;
}
static ExcelValue referencing_e72_variable;
ExcelValue referencing_e72() { if(variable_set[218] == 1) { return referencing_e72_variable; } else { return referencing_e72_default(); } }
void set_referencing_e72(ExcelValue newValue) { variable_set[218] = 1; referencing_e72_variable = newValue; }

ExcelValue referencing_f72() { return common7(); }
ExcelValue referencing_g72() { return common7(); }
ExcelValue tables_a1() {
  static ExcelValue result;
  if(variable_set[221] == 1) { return result;}
  result = ZERO;
  variable_set[221] = 1;
  return result;
}

ExcelValue tables_b2_default() {
  return constant100;
}
static ExcelValue tables_b2_variable;
ExcelValue tables_b2() { if(variable_set[222] == 1) { return tables_b2_variable; } else { return tables_b2_default(); } }
void set_tables_b2(ExcelValue newValue) { variable_set[222] = 1; tables_b2_variable = newValue; }

ExcelValue tables_c2_default() {
  return constant101;
}
static ExcelValue tables_c2_variable;
ExcelValue tables_c2() { if(variable_set[223] == 1) { return tables_c2_variable; } else { return tables_c2_default(); } }
void set_tables_c2(ExcelValue newValue) { variable_set[223] = 1; tables_c2_variable = newValue; }

ExcelValue tables_d2_default() {
  return constant102;
}
static ExcelValue tables_d2_variable;
ExcelValue tables_d2() { if(variable_set[224] == 1) { return tables_d2_variable; } else { return tables_d2_default(); } }
void set_tables_d2(ExcelValue newValue) { variable_set[224] = 1; tables_d2_variable = newValue; }

ExcelValue tables_b3_default() {
  return constant103;
}
static ExcelValue tables_b3_variable;
ExcelValue tables_b3() { if(variable_set[225] == 1) { return tables_b3_variable; } else { return tables_b3_default(); } }
void set_tables_b3(ExcelValue newValue) { variable_set[225] = 1; tables_b3_variable = newValue; }

ExcelValue tables_c3_default() {
  return constant103;
}
static ExcelValue tables_c3_variable;
ExcelValue tables_c3() { if(variable_set[226] == 1) { return tables_c3_variable; } else { return tables_c3_default(); } }
void set_tables_c3(ExcelValue newValue) { variable_set[226] = 1; tables_c3_variable = newValue; }

ExcelValue tables_d3() {
  static ExcelValue result;
  if(variable_set[227] == 1) { return result;}
  ExcelValue array0[] = {tables_b3(),tables_c3()};
  result = string_join(2, array0);
  variable_set[227] = 1;
  return result;
}

ExcelValue tables_b4_default() {
  return constant104;
}
static ExcelValue tables_b4_variable;
ExcelValue tables_b4() { if(variable_set[228] == 1) { return tables_b4_variable; } else { return tables_b4_default(); } }
void set_tables_b4(ExcelValue newValue) { variable_set[228] = 1; tables_b4_variable = newValue; }

ExcelValue tables_c4_default() {
  return constant105;
}
static ExcelValue tables_c4_variable;
ExcelValue tables_c4() { if(variable_set[229] == 1) { return tables_c4_variable; } else { return tables_c4_default(); } }
void set_tables_c4(ExcelValue newValue) { variable_set[229] = 1; tables_c4_variable = newValue; }

ExcelValue tables_d4() {
  static ExcelValue result;
  if(variable_set[230] == 1) { return result;}
  ExcelValue array0[] = {tables_b4(),tables_c4()};
  result = string_join(2, array0);
  variable_set[230] = 1;
  return result;
}

ExcelValue tables_f4() {
  static ExcelValue result;
  if(variable_set[231] == 1) { return result;}
  result = tables_c4();
  variable_set[231] = 1;
  return result;
}

ExcelValue tables_g4() {
  static ExcelValue result;
  if(variable_set[232] == 1) { return result;}
  static ExcelValue array0[3];
  array0[0] = tables_b4();
  array0[1] = tables_c4();
  array0[2] = tables_d4();
  ExcelValue array0_ev = EXCEL_RANGE(array0,1,3);
  result = excel_match(constant106,array0_ev,FALSE);
  variable_set[232] = 1;
  return result;
}

ExcelValue tables_h4() {
  static ExcelValue result;
  if(variable_set[233] == 1) { return result;}
  static ExcelValue array0[2];
  array0[0] = tables_c4();
  array0[1] = tables_d4();
  ExcelValue array0_ev = EXCEL_RANGE(array0,1,2);
  result = excel_match_2(constant105,array0_ev);
  variable_set[233] = 1;
  return result;
}

ExcelValue tables_b5() { return common9(); }
ExcelValue tables_c5() {
  static ExcelValue result;
  if(variable_set[235] == 1) { return result;}
  ExcelValue array0[] = {tables_c3(),tables_c4()};
  result = sum(2, array0);
  variable_set[235] = 1;
  return result;
}

ExcelValue tables_e6() {
  static ExcelValue result;
  if(variable_set[236] == 1) { return result;}
  result = tables_b2();
  variable_set[236] = 1;
  return result;
}

ExcelValue tables_f6() {
  static ExcelValue result;
  if(variable_set[237] == 1) { return result;}
  result = tables_c2();
  variable_set[237] = 1;
  return result;
}

ExcelValue tables_g6() {
  static ExcelValue result;
  if(variable_set[238] == 1) { return result;}
  result = tables_d2();
  variable_set[238] = 1;
  return result;
}

ExcelValue tables_e7() {
  static ExcelValue result;
  if(variable_set[239] == 1) { return result;}
  result = tables_b5();
  variable_set[239] = 1;
  return result;
}

ExcelValue tables_f7() {
  static ExcelValue result;
  if(variable_set[240] == 1) { return result;}
  result = tables_c5();
  variable_set[240] = 1;
  return result;
}

ExcelValue tables_g7() {
  static ExcelValue result;
  if(variable_set[241] == 1) { return result;}
  result = ZERO;
  variable_set[241] = 1;
  return result;
}

ExcelValue tables_e8() {
  static ExcelValue result;
  if(variable_set[242] == 1) { return result;}
  result = tables_b2();
  variable_set[242] = 1;
  return result;
}

ExcelValue tables_f8() {
  static ExcelValue result;
  if(variable_set[243] == 1) { return result;}
  result = tables_c2();
  variable_set[243] = 1;
  return result;
}

ExcelValue tables_g8() {
  static ExcelValue result;
  if(variable_set[244] == 1) { return result;}
  result = tables_d2();
  variable_set[244] = 1;
  return result;
}

ExcelValue tables_e9() {
  static ExcelValue result;
  if(variable_set[245] == 1) { return result;}
  result = tables_b3();
  variable_set[245] = 1;
  return result;
}

ExcelValue tables_f9() {
  static ExcelValue result;
  if(variable_set[246] == 1) { return result;}
  result = tables_c3();
  variable_set[246] = 1;
  return result;
}

ExcelValue tables_g9() {
  static ExcelValue result;
  if(variable_set[247] == 1) { return result;}
  result = tables_d3();
  variable_set[247] = 1;
  return result;
}

ExcelValue tables_c10() { return common2(); }
ExcelValue tables_e10() {
  static ExcelValue result;
  if(variable_set[249] == 1) { return result;}
  result = tables_b4();
  variable_set[249] = 1;
  return result;
}

ExcelValue tables_f10() {
  static ExcelValue result;
  if(variable_set[250] == 1) { return result;}
  result = tables_c4();
  variable_set[250] = 1;
  return result;
}

ExcelValue tables_g10() {
  static ExcelValue result;
  if(variable_set[251] == 1) { return result;}
  result = tables_d4();
  variable_set[251] = 1;
  return result;
}

ExcelValue tables_c11() { return common9(); }
ExcelValue tables_e11() {
  static ExcelValue result;
  if(variable_set[253] == 1) { return result;}
  result = tables_b5();
  variable_set[253] = 1;
  return result;
}

ExcelValue tables_f11() {
  static ExcelValue result;
  if(variable_set[254] == 1) { return result;}
  result = tables_c5();
  variable_set[254] = 1;
  return result;
}

ExcelValue tables_g11() {
  static ExcelValue result;
  if(variable_set[255] == 1) { return result;}
  result = ZERO;
  variable_set[255] = 1;
  return result;
}

ExcelValue tables_c12() {
  static ExcelValue result;
  if(variable_set[256] == 1) { return result;}
  result = tables_b5();
  variable_set[256] = 1;
  return result;
}

ExcelValue tables_c13() { return common10(); }
ExcelValue tables_c14() { return common10(); }
ExcelValue tables_g17() {
  static ExcelValue result;
  if(variable_set[259] == 1) { return result;}
  ExcelValue condition0 = more_than(tables_c4(),tables_d4());
  ExcelValue ifresult0;
  switch(condition0.type) {
  case ExcelBoolean:
    if(condition0.number == true) {
      ifresult0 = TRUE;
    } else {
      ifresult0 = FALSE;
    }
    break;
  case ExcelNumber:
    if(condition0.number == 0) {
      ifresult0 = FALSE;
    } else {
      ifresult0 = TRUE;
    }
    break;
  case ExcelEmpty: 
    ifresult0 = FALSE;
    break;
  case ExcelString:
  case ExcelRange:
    ifresult0 = VALUE;
    break;
  case ExcelError:
    ifresult0 = condition0;
    break;
  }
  result = ifresult0;
  variable_set[259] = 1;
  return result;
}

ExcelValue s_innapropriate_sheet_name__c4() {
  static ExcelValue result;
  if(variable_set[260] == 1) { return result;}
  result = valuetypes_a3();
  variable_set[260] = 1;
  return result;
}

ExcelValue ranges_f1() { return BLANK; }
ExcelValue ranges_f2() { return BLANK; }
ExcelValue ranges_f3() { return BLANK; }
ExcelValue ranges_a5() { return BLANK; }
ExcelValue ranges_b5() { return BLANK; }
ExcelValue ranges_c5() { return BLANK; }
ExcelValue ranges_d5() { return BLANK; }
ExcelValue referencing_f35() { return BLANK; }
ExcelValue referencing_g35() { return BLANK; }
ExcelValue referencing_h35() { return BLANK; }
ExcelValue referencing_i35() { return BLANK; }
ExcelValue referencing_k35() { return BLANK; }
ExcelValue referencing_l35() { return BLANK; }
ExcelValue referencing_f36() { return BLANK; }
ExcelValue referencing_g36() { return BLANK; }
ExcelValue referencing_h36() { return BLANK; }
ExcelValue referencing_i36() { return BLANK; }
ExcelValue referencing_k36() { return BLANK; }
ExcelValue referencing_l36() { return BLANK; }
ExcelValue referencing_g37() { return BLANK; }
ExcelValue referencing_h37() { return BLANK; }
ExcelValue referencing_i37() { return BLANK; }
ExcelValue referencing_j37() { return BLANK; }
ExcelValue referencing_k37() { return BLANK; }
ExcelValue referencing_l37() { return BLANK; }
ExcelValue referencing_f38() { return BLANK; }
ExcelValue referencing_g38() { return BLANK; }
ExcelValue referencing_h38() { return BLANK; }
ExcelValue referencing_j38() { return BLANK; }
ExcelValue referencing_k38() { return BLANK; }
ExcelValue referencing_l38() { return BLANK; }
ExcelValue referencing_f39() { return BLANK; }
ExcelValue referencing_g39() { return BLANK; }
ExcelValue referencing_i39() { return BLANK; }
ExcelValue referencing_j39() { return BLANK; }
ExcelValue referencing_k39() { return BLANK; }
ExcelValue referencing_l39() { return BLANK; }
ExcelValue referencing_f40() { return BLANK; }
ExcelValue referencing_h40() { return BLANK; }
ExcelValue referencing_i40() { return BLANK; }
ExcelValue referencing_k40() { return BLANK; }
ExcelValue referencing_l40() { return BLANK; }
ExcelValue referencing_f41() { return BLANK; }
ExcelValue referencing_h41() { return BLANK; }
ExcelValue referencing_i41() { return BLANK; }
ExcelValue referencing_k41() { return BLANK; }
ExcelValue referencing_l41() { return BLANK; }
ExcelValue referencing_g42() { return BLANK; }
ExcelValue referencing_h42() { return BLANK; }
ExcelValue referencing_i42() { return BLANK; }
ExcelValue referencing_j42() { return BLANK; }
ExcelValue referencing_k42() { return BLANK; }
ExcelValue referencing_n42() { return BLANK; }
ExcelValue referencing_g43() { return BLANK; }
ExcelValue referencing_h43() { return BLANK; }
ExcelValue referencing_i43() { return BLANK; }
ExcelValue referencing_j43() { return BLANK; }
ExcelValue referencing_k43() { return BLANK; }
ExcelValue referencing_n43() { return BLANK; }
ExcelValue referencing_f44() { return BLANK; }
ExcelValue referencing_g44() { return BLANK; }
ExcelValue referencing_h44() { return BLANK; }
ExcelValue referencing_i44() { return BLANK; }
ExcelValue referencing_j44() { return BLANK; }
ExcelValue referencing_k44() { return BLANK; }
ExcelValue referencing_f45() { return BLANK; }
ExcelValue referencing_h45() { return BLANK; }
ExcelValue referencing_i45() { return BLANK; }
ExcelValue referencing_k45() { return BLANK; }
ExcelValue referencing_l45() { return BLANK; }
ExcelValue referencing_f46() { return BLANK; }
ExcelValue referencing_i46() { return BLANK; }
ExcelValue referencing_j46() { return BLANK; }
ExcelValue referencing_k46() { return BLANK; }
ExcelValue referencing_l46() { return BLANK; }
ExcelValue referencing_f47() { return BLANK; }
ExcelValue referencing_g47() { return BLANK; }
ExcelValue referencing_h47() { return BLANK; }
ExcelValue referencing_i47() { return BLANK; }
ExcelValue referencing_j47() { return BLANK; }
ExcelValue referencing_l47() { return BLANK; }
ExcelValue common0() {
  static ExcelValue result;
  if(variable_set[342] == 1) { return result;}
  result = add(add(referencing_a4(),ONE),ONE);
  variable_set[342] = 1;
  return result;
}

ExcelValue common1() {
  static ExcelValue result;
  if(variable_set[343] == 1) { return result;}
  result = add(referencing_a4(),ONE);
  variable_set[343] = 1;
  return result;
}

ExcelValue common2() {
  static ExcelValue result;
  if(variable_set[344] == 1) { return result;}
  ExcelValue array0[] = {tables_b5(),tables_c5()};
  result = sum(2, array0);
  variable_set[344] = 1;
  return result;
}

ExcelValue common3() {
  static ExcelValue result;
  if(variable_set[345] == 1) { return result;}
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
  ExcelValue array0_ev = EXCEL_RANGE(array0,4,4);
  static ExcelValue array1[4];
  array1[0] = referencing_c15();
  array1[1] = referencing_d15();
  array1[2] = referencing_e15();
  array1[3] = referencing_f15();
  ExcelValue array1_ev = EXCEL_RANGE(array1,1,4);
  result = excel_index(array0_ev,BLANK,excel_match(referencing_c22(),array1_ev,ZERO));
  variable_set[345] = 1;
  return result;
}

ExcelValue common7() {
  static ExcelValue result;
  if(variable_set[346] == 1) { return result;}
  static ExcelValue array1[3];
  array1[0] = multiply(excel_equal(referencing_e72(),referencing_e68()),referencing_f68());
  array1[1] = multiply(excel_equal(referencing_e72(),referencing_e69()),referencing_f69());
  array1[2] = multiply(excel_equal(referencing_e72(),referencing_e70()),referencing_f70());
  ExcelValue array1_ev = EXCEL_RANGE(array1,3,1);
  ExcelValue array0[] = {array1_ev};
  result = sum(1, array0);
  variable_set[346] = 1;
  return result;
}

ExcelValue common9() {
  static ExcelValue result;
  if(variable_set[347] == 1) { return result;}
  ExcelValue array0[] = {tables_b3(),tables_b4()};
  result = sum(2, array0);
  variable_set[347] = 1;
  return result;
}

ExcelValue common10() {
  static ExcelValue result;
  if(variable_set[348] == 1) { return result;}
  ExcelValue array0[] = {tables_b3(),tables_c3(),tables_d3(),tables_b4(),tables_c4(),tables_d4()};
  result = sum(6, array0);
  variable_set[348] = 1;
  return result;
}

// Start of named references
// End of named references
