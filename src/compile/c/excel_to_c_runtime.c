#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>

// I predefine an array of ExcelValues to store calculations
// Probably bad practice. At the very least, I should make it
// link to the cell reference in some way.
#define MAX_EXCEL_VALUE_HEAP_SIZE 100000

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

// My little heap
ExcelValue cells[MAX_EXCEL_VALUE_HEAP_SIZE];
int cell_counter = 0;

// Clears the heap
void reset() {
	cell_counter = 0;
}

// The object initializers
ExcelValue new_excel_number(double number) {
	cell_counter++;
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelNumber;
	new_cell.number = number;
	return new_cell;
};

ExcelValue new_excel_string(char *string) {
	cell_counter++;
	ExcelValue new_cell = 	cells[cell_counter];
	new_cell.type = ExcelString;
	new_cell.string = string;
	return new_cell;
};

ExcelValue new_excel_range(void *array, int rows, int columns) {
	cell_counter++;
	ExcelValue new_cell = cells[cell_counter];
	new_cell.type = ExcelRange;
	new_cell.array =array;
	new_cell.rows = rows;
	new_cell.columns = columns;
	return new_cell;
};

// Constants
ExcelValue BLANK = {.type = ExcelEmpty };

// Booleans
ExcelValue TRUE = {.type = ExcelBoolean, .number = true };
ExcelValue FALSE = {.type = ExcelBoolean, .number = false };

// Errors
ExcelValue VALUE = {.type = ExcelError, .number = 0};
ExcelValue NAME = {.type = ExcelError, .number = 1};
ExcelValue DIV0 = {.type = ExcelError, .number = 2};
ExcelValue REF = {.type = ExcelError, .number = 3};
ExcelValue NA = {.type = ExcelError, .number = 4};

// This is the error flag
int conversion_error = 0;

// Extracts numbers from ExcelValues
// Excel treats empty cells as zero
double number_from(ExcelValue v) {
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
	
ExcelValue excel_abs(ExcelValue a_v) {
	NUMBER(a_v, a)
	CHECK_FOR_CONVERSION_ERROR
	
	if(a >= 0.0 ) {
		return a_v;
	} else {
		return new_excel_number(-a);
	}
}

ExcelValue add(ExcelValue a_v, ExcelValue b_v) {
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a + b);
}

ExcelValue excel_and(int array_size, ExcelValue *array) {
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
	
struct average_result calculate_average(int array_size, ExcelValue *array) {
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
	return r;
}

ExcelValue average(int array_size, ExcelValue *array) {
	struct average_result r = calculate_average(array_size, array);
	if(r.has_error == true) return r.error;
	if(r.count == 0) return DIV0;
	return new_excel_number(r.sum/r.count);
}

ExcelValue choose(ExcelValue index_v, int array_size, ExcelValue *array) {
	if(index_v.type == ExcelError) return index_v;
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

ExcelValue count(int array_size, ExcelValue *array) {
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

ExcelValue counta(int array_size, ExcelValue *array) {
	int i;
	int n = 0;
	ExcelValue current_excel_value;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		if(current_excel_value.type != ExcelEmpty) n++;
	 }
	 return new_excel_number(n);
}

ExcelValue divide(ExcelValue a_v, ExcelValue b_v) {
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	if(b == 0) return DIV0;
	return new_excel_number(a / b);
}
	
ExcelValue subtract(ExcelValue a_v, ExcelValue b_v) {
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a - b);
}

ExcelValue multiply(ExcelValue a_v, ExcelValue b_v) {
	NUMBER(a_v, a)
	NUMBER(b_v, b)
	CHECK_FOR_CONVERSION_ERROR
	return new_excel_number(a * b);
}


ExcelValue sum(int array_size, ExcelValue *array) {
	double total = 0;
	int i;
	double number;
	ExcelValue current_excel_value;
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		if(current_excel_value.type == ExcelRange) {
			number = number_from(sum( current_excel_value.rows * current_excel_value.columns, current_excel_value.array ));
		} else {
			number = number_from(current_excel_value);					
		}
		CHECK_FOR_CONVERSION_ERROR
		total += number;
	}
	return new_excel_number(total);
}

int main()
{
	// Test ABS
	assert(excel_abs(new_excel_number(1)).number == 1);
	assert(excel_abs(new_excel_number(-1)).number == 1);
	assert(excel_abs(VALUE).type == ExcelError);
	
	// Test ADD
	assert(add(new_excel_number(1),new_excel_number(-2.5)).number == -1.5);
	assert(add(new_excel_number(1),VALUE).type == ExcelError);
	
	// Test AND
	ExcelValue true_array1[] = { TRUE, new_excel_number(10)};
	ExcelValue true_array2[] = { new_excel_number(1) };
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
	assert(choose(new_excel_number(1),4,array1).number == 10);
	assert(choose(new_excel_number(4),4,array1).type == ExcelBoolean);
	assert(choose(new_excel_number(0),4,array1).type == ExcelError);
	assert(choose(new_excel_number(5),4,array1).type == ExcelError);
	assert(choose(new_excel_number(1),4,array3).type == ExcelError);	
	
	// Test COUNT
	assert(count(4,array1).number == 2);
	assert(count(3,array2).number == 3);
	assert(count(4,array3).number == 3);
	
	// Test COUNTA
	ExcelValue count_a_test_array_1[] = { new_excel_number(10), new_excel_number(5), TRUE, FALSE, new_excel_string("Hello"), VALUE, BLANK};
	assert(counta(7, count_a_test_array_1).number == 6);
	
	// Test divide
	assert(divide(new_excel_number(12.4),new_excel_number(3.2)).number == 3.875);
	assert(divide(new_excel_number(12.4),new_excel_number(0)).type == ExcelError);
	
	// // Test number handling
	// ExcelValue one = new_excel_number(38.8);
	// assert(one.number == 38.8);
	// assert(one.type == ExcelNumber);
	// 
	// // Test string handling
	// char *string = "Hello world";
	// ExcelValue two = new_excel_string("Hello world");
	// ExcelValue three = new_excel_string("Bye");
	// assert(strcmp(two.string,string) == 0);
	// assert(strcmp(a1().string,string) == 0);
	// assert(strcmp(three.string,"Bye") == 0);
	// 
	// //printf("a3: %f",a3().number);
	// assert(a3().number == 24);
	// assert(a5().number == 7);
	// assert(a6().type == ExcelError);
	// 
	// assert(a7().type == ExcelRange);
	// assert(a7().rows == 1);
	// assert(a7().columns == 3);
	// ExcelValue temp = a7();
	// ExcelValue *p = temp.array;
	// assert(p[0].type == ExcelString);
	// assert(p[1].type == ExcelNumber);
	
	return 0;
}