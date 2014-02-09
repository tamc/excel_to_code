#include "excel_to_c_runtime.c"

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
  // Should deal with very long string joins
	ExcelValue string_join_array_6[] = {new_excel_string("0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"), new_excel_string("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")};
  assert(string_join(2, string_join_array_6).string[0] == '0');
  free_all_allocated_memory();
	
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

  // Test the EXP function
  assert(excel_exp(BLANK).number == 1);
  assert(excel_exp(ZERO).number == 1);
  assert(excel_exp(ONE).number == 2.718281828459045);
  assert(excel_exp(new_excel_string("1")).number == 2.718281828459045);
  assert(excel_exp(FALSE).number == 1);
  assert(excel_exp(TRUE).number == 2.718281828459045);
  assert(excel_exp(DIV0).type == ExcelError);

  // Test the ISBLANK function
  assert(excel_isblank(BLANK).type == ExcelBoolean);
  assert(excel_isblank(BLANK).number == true);
  assert(excel_isblank(ZERO).type == ExcelBoolean);
  assert(excel_isblank(ZERO).number == false);
  assert(excel_isblank(TRUE).number == false);
  assert(excel_isblank(FALSE).number == false);
  assert(excel_isblank(new_excel_string("")).number == false);
  
  // Test AVERAGEIFS function
  ExcelValue averageifs_array_1[] = {new_excel_number(10),new_excel_number(100),BLANK};
  ExcelValue averageifs_array_1_v = new_excel_range(averageifs_array_1,3,1);
  ExcelValue averageifs_array_2[] = {new_excel_string("pear"),new_excel_string("bear"),new_excel_string("apple")};
  ExcelValue averageifs_array_2_v = new_excel_range(averageifs_array_2,3,1);
  ExcelValue averageifs_array_3[] = {new_excel_number(1),new_excel_number(2),new_excel_number(3),new_excel_number(4),new_excel_number(5),new_excel_number(5)};
  ExcelValue averageifs_array_3_v = new_excel_range(averageifs_array_3,6,1);
  ExcelValue averageifs_array_4[] = {new_excel_string("CO2"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CH4"),new_excel_string("N2O"),new_excel_string("CO2")};
  ExcelValue averageifs_array_4_v = new_excel_range(averageifs_array_4,6,1);
  ExcelValue averageifs_array_5[] = {new_excel_string("1A"),new_excel_string("1A"),new_excel_string("1A"),new_excel_number(4),new_excel_number(4),new_excel_number(5)};
  ExcelValue averageifs_array_5_v = new_excel_range(averageifs_array_5,6,1);
  
  // ... should only average values that meet all of the criteria
  ExcelValue averageifs_array_6[] = { averageifs_array_1_v, new_excel_number(10), averageifs_array_2_v, new_excel_string("Bear") };
  assert(averageifs(averageifs_array_1_v,4,averageifs_array_6).type == ExcelError);
  
  ExcelValue averageifs_array_7[] = { averageifs_array_1_v, new_excel_number(10), averageifs_array_2_v, new_excel_string("Pear") };
  assert(averageifs(averageifs_array_1_v,4,averageifs_array_7).number == 10.0);
  
  // ... should work when single cells are given where ranges expected
  ExcelValue averageifs_array_8[] = { new_excel_string("CAR"), new_excel_string("CAR"), new_excel_string("FCV"), new_excel_string("FCV")};
  assert(averageifs(new_excel_number(0.143897265452564), 4, averageifs_array_8).number == 0.143897265452564);

  // ... should match numbers with strings that contain numbers
  ExcelValue averageifs_array_9[] = { new_excel_number(10), new_excel_string("10.0")};
  assert(averageifs(new_excel_number(100),2,averageifs_array_9).number == 100);
  
  ExcelValue averageifs_array_10[] = { averageifs_array_4_v, new_excel_string("CO2"), averageifs_array_5_v, new_excel_number(2)};
  assert(averageifs(averageifs_array_3_v,4, averageifs_array_10).type == ExcelError);
  
  // ... should match with strings that contain criteria
  ExcelValue averageifs_array_10a[] = { averageifs_array_3_v, new_excel_string("=5")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10a).number == 5);

  ExcelValue averageifs_array_10b[] = { averageifs_array_3_v, new_excel_string("<>3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10b).number == 3.4);

  ExcelValue averageifs_array_10c[] = { averageifs_array_3_v, new_excel_string("<3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10c).number == 1.5);
  
  ExcelValue averageifs_array_10d[] = { averageifs_array_3_v, new_excel_string("<=3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10d).number == 2);

  ExcelValue averageifs_array_10e[] = { averageifs_array_3_v, new_excel_string(">3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10e).number == 14.0/3.0);

  ExcelValue averageifs_array_10f[] = { averageifs_array_3_v, new_excel_string(">=3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10f).number == (3.0+4.0+5.0+5.0)/4.0);
  
  // ... should treat BLANK as an empty string when in the check_range, but not in the criteria
  ExcelValue averageifs_array_11[] = { BLANK, new_excel_number(20)};
  assert(averageifs(new_excel_number(100),2,averageifs_array_11).type == ExcelError);
  
  ExcelValue averageifs_array_12[] = {BLANK, new_excel_string("")};
  assert(averageifs(new_excel_number(100),2,averageifs_array_12).number == 100);
  
  ExcelValue averageifs_array_13[] = {BLANK, BLANK};
  assert(averageifs(new_excel_number(100),2,averageifs_array_13).type == ExcelError);
    
  // ... should return an error if range argument is an error
  assert(averageifs(REF,2,averageifs_array_13).type == ExcelError);

  // Tests for the FORECAST function
  ExcelValue forecast_array1[] = { new_excel_number(1), new_excel_number(2), new_excel_number(3), new_excel_number(4), new_excel_number(5)};
  ExcelValue forecast_array2[] = { new_excel_number(2), new_excel_number(3), new_excel_number(4), new_excel_number(5), new_excel_number(6)};
  ExcelValue forecast_array1_v = new_excel_range(forecast_array1,5,1);
  ExcelValue forecast_array2_v = new_excel_range(forecast_array2,5,1);

  assert(forecast(new_excel_number(0), forecast_array2_v, forecast_array1_v).number == 1);
  assert(forecast(new_excel_number(1), forecast_array2_v, forecast_array1_v).number == 2);
  assert(forecast(new_excel_number(6), forecast_array2_v, forecast_array1_v).number == 7);

  ExcelValue forecast_array3[] = { BLANK, new_excel_number(2), new_excel_number(3), new_excel_number(4), BLANK};
  ExcelValue forecast_array3_v = new_excel_range(forecast_array3,5,1);

  assert(forecast(new_excel_number(6), forecast_array2_v, forecast_array3_v).number == 7);

  // Tests ENSURE_IS_NUMBER function
  assert(ensure_is_number(new_excel_number(1.3)).type == ExcelNumber);
  assert(ensure_is_number(new_excel_number(1.3)).number == 1.3);
  assert(ensure_is_number(BLANK).type == ExcelNumber);
  assert(ensure_is_number(BLANK).number == 0);
  assert(ensure_is_number(TRUE).type == ExcelNumber);
  assert(ensure_is_number(TRUE).number == 1.0);
  assert(ensure_is_number(FALSE).type == ExcelNumber);
  assert(ensure_is_number(FALSE).number == 0.0);
  assert(ensure_is_number(new_excel_string("1.3")).type == ExcelNumber);
  assert(ensure_is_number(new_excel_string("1.3")).number == 1.3);
  assert(ensure_is_number(new_excel_string("BASDASD")).type == ExcelError);
  assert(ensure_is_number(DIV0).type == ExcelError);

  // Release memory
  free_all_allocated_memory();
  
  // Yay!
  printf("All tests passed\n");
  
  return 0;
}

int main() {
	return test_functions();
}
