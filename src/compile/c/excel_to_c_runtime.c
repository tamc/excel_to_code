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
				if(x.type == ExcelEmpty) x = ZERO;
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
				if(x.type == ExcelEmpty) x = ZERO;
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

  if(strcasecmp(type_v.string, "s") == 0 ) {
    return scurve(currentYear_v, startValue_v, endValue_v, duration_v, startYear_v);
  }

  if(strcasecmp(type_v.string, "hs") == 0 ) {
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
