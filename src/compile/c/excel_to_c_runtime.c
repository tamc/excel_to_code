#include <stdio.h>
#include <assert.h>
#include <string.h>
#include <stdlib.h>
#include <ctype.h>

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

ExcelValue cells[100000];
int cell_counter = 0;
int conversion_error = 0;

void reset() {
	cell_counter = 0;
	conversion_error = 0;
}

ExcelValue new_excel_number(double number) {
	cell_counter++;
	cells[cell_counter].type = ExcelNumber;
	cells[cell_counter].number = number;
	return cells[cell_counter];
};

ExcelValue new_excel_string(char *string) {
	cell_counter++;
	cells[cell_counter].type = ExcelString;
	cells[cell_counter].string = string;
	return cells[cell_counter];
};

ExcelValue new_excel_range(void *array, int rows, int columns) {
	cell_counter++;
	cells[cell_counter].type = ExcelRange;
	cells[cell_counter].array =array;
	cells[cell_counter].rows = rows;
	cells[cell_counter].columns = columns;
	return cells[cell_counter];
};

// Constants
ExcelValue BLANK = {.type = ExcelEmpty };

ExcelValue TRUE = {.type = ExcelBoolean, .number = 1 };
ExcelValue FALSE = {.type = ExcelBoolean, .number = 0 };

ExcelValue VALUE = {.type = ExcelError, .number = 0};
ExcelValue NAME = {.type = ExcelError, .number = 1};
ExcelValue DIV0 = {.type = ExcelError, .number = 2};
ExcelValue REF = {.type = ExcelError, .number = 3};
ExcelValue NA = {.type = ExcelError, .number = 4};

double number_from(ExcelValue v) {
	char *s;
	char * p;
	double n;
	ExcelValue *array;
	switch (v.type) {
  	  case ExcelNumber: return v.number;
	  case ExcelEmpty: return 0;
	  case ExcelBoolean: return v.number;
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
	  case ExcelError: return 0;
  }
  return 0;
}

ExcelValue excel_abs(ExcelValue a_v) {
	double a;
	
	if(a_v.type == ExcelError) {
		return a_v;
	}
	
	a = number_from(a_v);
	
	if(conversion_error) {
		conversion_error = 0;
		return VALUE;
	}
	
	if(a >= 0.0 ) {
		return a_v;
	} else {
		return new_excel_number(-a);
	}
}

ExcelValue add(ExcelValue a_v, ExcelValue b_v) {
	double a, b;

	if(a_v.type == ExcelError) {
		return a_v;
	}

	if(b_v.type == ExcelError) {
		return b_v;
	}

	a = number_from(a_v);
	b = number_from(b_v);

	if(conversion_error) {
		conversion_error = 0;
		return VALUE;
	}

	return new_excel_number(a + b);
}

ExcelValue excel_and(int array_size, ExcelValue *array) {
	int i;
	ExcelValue current_excel_value, array_result;
	
	for(i=0;i<array_size;i++) {
		current_excel_value = array[i];
		switch (current_excel_value.type) {
	  	  case ExcelNumber: 
			  if(current_excel_value.number == 0) {
				  return FALSE;
			  }
			  break;
		  case ExcelBoolean: 
			  if(current_excel_value.number == 0) {
				  return FALSE;
			  }
			  break;
		  case ExcelRange: 
		  	array_result = excel_and( current_excel_value.rows * current_excel_value.columns, current_excel_value.array );
			if(array_result.type == ExcelError) {
				return array_result;
			}
			if(array_result.type == ExcelBoolean) {
				if(array_result.number == 0) {
					return FALSE;
				}
			}
			break;
		  case ExcelString:
			 break;
		  case ExcelEmpty:
			 break;
		  case ExcelError:
			 return current_excel_value;
			 break;
		 }
	 }
	 return TRUE;
}
	
ExcelValue subtract(ExcelValue a_v, ExcelValue b_v) {
	double a, b;

	if(a_v.type == ExcelError) {
		return a_v;
	}

	if(b_v.type == ExcelError) {
		return b_v;
	}

	a = number_from(a_v);
	b = number_from(b_v);

	if(conversion_error) {
		conversion_error = 0;
		return VALUE;
	}

	return new_excel_number(a - b);
}

ExcelValue multiply(ExcelValue a_v, ExcelValue b_v) {
	double a, b;

	if(a_v.type == ExcelError) {
		return a_v;
	}

	if(b_v.type == ExcelError) {
		return b_v;
	}

	a = number_from(a_v);
	b = number_from(b_v);

	if(conversion_error) {
		conversion_error = 0;
		return VALUE;
	}

	return new_excel_number(a * b);
}

ExcelValue divide(ExcelValue a_v, ExcelValue b_v) {
	double a, b;

	if(a_v.type == ExcelError) {
		return a_v;
	}

	if(b_v.type == ExcelError) {
		return b_v;
	}

	a = number_from(a_v);
	b = number_from(b_v);

	if(conversion_error) {
		conversion_error = 0;
		return VALUE;
	}
	
	if(b == 0) {
		return DIV0;
	}

	return new_excel_number(a + b);
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
		if(conversion_error) {
			conversion_error = 0;
			return VALUE;
		}
		total += number;
	}
	return new_excel_number(total);
}

int main()
{
	// Test abs
	assert(excel_abs(new_excel_number(1)).number == 1);
	assert(excel_abs(new_excel_number(-1)).number == 1);
	assert(excel_abs(VALUE).type == ExcelError);
	
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