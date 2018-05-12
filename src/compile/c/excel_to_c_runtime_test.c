#include "excel_to_c_runtime.c"

int test_functions() {
  // Assertions
  assert_equal(NA, NA, "NA == NA");
  assert_equal(ZERO, EXCEL_NUMBER(0), "ZERO == ZERO");
  assert_equal(EXCEL_NUMBER(-0.0), EXCEL_NUMBER(0.0), "Negative ZERO == ZERO");
  assert_equal(EXCEL_NUMBER(0.0), EXCEL_NUMBER(-0.0), "ZERO == negative ZERO");
  assert_equal(EXCEL_NUMBER(0.0), EXCEL_NUMBER(1e-10), "Expected zero, got almost zero");
  assert_equal(EXCEL_NUMBER(0.0), EXCEL_NUMBER(-1e-10), "Expected zero, got negative almost zero");
  assert_equal(EXCEL_NUMBER(-0.0), EXCEL_NUMBER(1e-10), "Expected negative zero, got almost zero");
  assert_equal(EXCEL_NUMBER(-0.0), EXCEL_NUMBER(-1e-10), "Expected negative zero, got negative almost zero");

  // Test ABS
  assert(excel_abs(ONE).number == 1);
  assert(excel_abs(EXCEL_NUMBER(-1)).number == 1);
  assert(excel_abs(VALUE).type == ExcelError);

  // Test ADD
  assert(add(ONE,EXCEL_NUMBER(-2.5)).number == -1.5);
  assert(add(ONE,VALUE).type == ExcelError);

  // Test AND
  ExcelValue true_array1[] = { TRUE, EXCEL_NUMBER(10)};
  ExcelValue true_array2[] = { ONE };
  ExcelValue false_array1[] = { FALSE, EXCEL_NUMBER(10)};
  ExcelValue false_array2[] = { TRUE, EXCEL_NUMBER(0)};
  // ExcelValue error_array1[] = { EXCEL_NUMBER(10)}; // Not implemented
  ExcelValue error_array2[] = { TRUE, NA};
  assert(excel_and(2,true_array1).number == 1);
  assert(excel_and(1,true_array2).number == 1);
  assert(excel_and(2,false_array1).number == 0);
  assert(excel_and(2,false_array2).number == 0);
  // assert(excel_and(1,error_array1).type == ExcelError); // Not implemented
  assert(excel_and(2,error_array2).type == ExcelError);

  // Test OR
  assert(excel_or(2,true_array1).number == 1);
  assert(excel_or(1,true_array2).number == 1);
  assert(excel_or(2,false_array1).number == 0);
  assert(excel_or(2,false_array2).number == 1);
  //assert(excel_or(2,error_array2).type == ExcelError); // Not implemented

  // Test NOT
  assert(excel_not(TRUE).number == false);
  assert(excel_not(FALSE).number == true);
  assert(excel_not(ZERO).number == true);
  assert(excel_not(ONE).number == false);
  assert(excel_not(TWO).number == false);
  assert(excel_not(BLANK).number == true);
  assert(excel_not(NA).type == ExcelError);
  assert(excel_not(EXCEL_STRING("hello world")).type == ExcelError);

  // Test AVERAGE
  ExcelValue array1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(5), TRUE, FALSE};
  ExcelValue array1_v = EXCEL_RANGE(array1,2,2);
  ExcelValue array2[] = { array1_v, EXCEL_NUMBER(9), EXCEL_STRING("Hello")};
  ExcelValue array3[] = { array1_v, EXCEL_NUMBER(9), EXCEL_STRING("Hello"), VALUE};
  assert(average(4, array1).number == 7.5);
  assert(average(3, array2).number == 8);
  assert(average(4, array3).type == ExcelError);

  // Test CHOOSE
  assert(choose(ONE,4,array1).number == 10);
  assert(choose(EXCEL_NUMBER(4),4,array1).type == ExcelBoolean);
  assert(choose(EXCEL_NUMBER(0),4,array1).type == ExcelError);
  assert(choose(EXCEL_NUMBER(5),4,array1).type == ExcelError);
  assert(choose(ONE,4,array3).type == ExcelError);

  // Test COUNT
  assert(count(4,array1).number == 2);
  assert(count(3,array2).number == 3);
  assert(count(4,array3).number == 3);

  // Test Large
  ExcelValue large_test_array_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), EXCEL_NUMBER(500), BLANK };
  ExcelValue large_test_array_1_v = EXCEL_RANGE(large_test_array_1, 1, 4);
  assert(large(large_test_array_1_v, EXCEL_NUMBER(1)).number == 500);
  assert(large(large_test_array_1_v, EXCEL_NUMBER(2)).number == 100);
  assert(large(large_test_array_1_v, EXCEL_NUMBER(3)).number == 10);
  assert(large(large_test_array_1_v, EXCEL_NUMBER(4)).type == ExcelError);
  assert(large(EXCEL_NUMBER(500), EXCEL_NUMBER(1)).number == 500);
  ExcelValue large_test_array_2[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), EXCEL_NUMBER(500), VALUE };
  ExcelValue large_test_array_2_v = EXCEL_RANGE(large_test_array_2, 1, 4);
  assert(large(large_test_array_2_v,EXCEL_NUMBER(2)).type == ExcelError);
  assert(large(EXCEL_NUMBER(500),VALUE).type == ExcelError);


  // Test COUNTA
  ExcelValue count_a_test_array_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(5), TRUE, FALSE, EXCEL_STRING("Hello"), VALUE, BLANK};
  ExcelValue count_a_test_array_1_v = EXCEL_RANGE(count_a_test_array_1,7,1);
  ExcelValue count_a_test_array_2[] = {EXCEL_STRING("Bye"),count_a_test_array_1_v};
  assert(counta(7, count_a_test_array_1).number == 6);
  assert(counta(2, count_a_test_array_2).number == 7);

  // Test divide
  assert(divide(EXCEL_NUMBER(12.4),EXCEL_NUMBER(3.2)).number == 3.875);
  assert(divide(EXCEL_NUMBER(12.4),EXCEL_NUMBER(0)).type == ExcelError);

  // Test excel_equal
  assert(excel_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(3.4)).type == ExcelBoolean);
  assert(excel_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(3.4)).number == false);
  assert(excel_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(1.2)).number == true);
  assert(excel_equal(EXCEL_STRING("hello"), EXCEL_STRING("HELLO")).number == true);
  assert(excel_equal(EXCEL_STRING("hello world"), EXCEL_STRING("HELLO")).number == false);
  assert(excel_equal(EXCEL_STRING("1"), ONE).number == false);
  assert(excel_equal(DIV0, ONE).type == ExcelError);

  // Test not_equal
  assert(not_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(3.4)).type == ExcelBoolean);
  assert(not_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(3.4)).number == true);
  assert(not_equal(EXCEL_NUMBER(1.2),EXCEL_NUMBER(1.2)).number == false);
  assert(not_equal(EXCEL_STRING("hello"), EXCEL_STRING("HELLO")).number == false);
  assert(not_equal(EXCEL_STRING("hello world"), EXCEL_STRING("HELLO")).number == true);
  assert(not_equal(EXCEL_STRING("1"), ONE).number == true);
  assert(not_equal(DIV0, ONE).type == ExcelError);

  // Test excel_if
  // Two argument version
  assert(excel_if_2(TRUE,EXCEL_NUMBER(10)).type == ExcelNumber);
  assert(excel_if_2(TRUE,EXCEL_NUMBER(10)).number == 10);
  assert(excel_if_2(FALSE,EXCEL_NUMBER(10)).type == ExcelBoolean);
  assert(excel_if_2(FALSE,EXCEL_NUMBER(10)).number == false);
  assert(excel_if_2(NA,EXCEL_NUMBER(10)).type == ExcelError);
  // Three argument version
  assert(excel_if(TRUE,EXCEL_NUMBER(10),EXCEL_NUMBER(20)).type == ExcelNumber);
  assert(excel_if(TRUE,EXCEL_NUMBER(10),EXCEL_NUMBER(20)).number == 10);
  assert(excel_if(FALSE,EXCEL_NUMBER(10),EXCEL_NUMBER(20)).type == ExcelNumber);
  assert(excel_if(FALSE,EXCEL_NUMBER(10),EXCEL_NUMBER(20)).number == 20);
  assert(excel_if(NA,EXCEL_NUMBER(10),EXCEL_NUMBER(20)).type == ExcelError);
  assert(excel_if(TRUE,EXCEL_NUMBER(10),NA).type == ExcelNumber);
  assert(excel_if(TRUE,EXCEL_NUMBER(10),NA).number == 10);

  // Test excel_match
  ExcelValue excel_match_array_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100) };
  ExcelValue excel_match_array_1_v = EXCEL_RANGE(excel_match_array_1,1,2);
  ExcelValue excel_match_array_2[] = { EXCEL_STRING("Pear"), EXCEL_STRING("Bear"), EXCEL_STRING("Apple") };
  ExcelValue excel_match_array_2_v = EXCEL_RANGE(excel_match_array_2,3,1);
  ExcelValue excel_match_array_4[] = { ONE, BLANK, EXCEL_NUMBER(0) };
  ExcelValue excel_match_array_4_v = EXCEL_RANGE(excel_match_array_4,1,3);
  ExcelValue excel_match_array_5[] = { ONE, EXCEL_NUMBER(0), BLANK };
  ExcelValue excel_match_array_5_v = EXCEL_RANGE(excel_match_array_5,1,3);
  ExcelValue excel_match_array_6[] = { EXCEL_STRING(""), ONE, TWO, THREE, FOUR };
  ExcelValue excel_match_array_6_v = EXCEL_RANGE(excel_match_array_6,5,1);

  // Two argument version
  assert(excel_match_2(EXCEL_NUMBER(14),excel_match_array_1_v).number == 1);
  assert(excel_match_2(EXCEL_NUMBER(110),excel_match_array_1_v).number == 2);
  assert(excel_match_2(EXCEL_NUMBER(-10),excel_match_array_1_v).type == ExcelError);

  // Three argument version
  assert(excel_match(EXCEL_NUMBER(10.0), excel_match_array_1_v, EXCEL_NUMBER(0) ).number == 1);
  assert(excel_match(EXCEL_NUMBER(100.0), excel_match_array_1_v, EXCEL_NUMBER(0) ).number == 2);
  assert(excel_match(EXCEL_NUMBER(1000.0), excel_match_array_1_v, EXCEL_NUMBER(0) ).type == ExcelError);
  assert(excel_match(EXCEL_STRING("bEAr"), excel_match_array_2_v, EXCEL_NUMBER(0) ).number == 2);
  assert(excel_match(EXCEL_NUMBER(1000.0), excel_match_array_1_v, ONE ).number == 2);
  assert(excel_match(EXCEL_NUMBER(1.0), excel_match_array_1_v, ONE ).type == ExcelError);
  assert(excel_match(EXCEL_STRING("Care"), excel_match_array_2_v, EXCEL_NUMBER(-1) ).number == 1  );
  assert(excel_match(EXCEL_STRING("Zebra"), excel_match_array_2_v, EXCEL_NUMBER(-1) ).type == ExcelError);
  assert(excel_match(EXCEL_STRING("a"), excel_match_array_2_v, EXCEL_NUMBER(-1) ).number == 2);
  // EMPTY STRINGS
  assert(excel_match(EXCEL_NUMBER(1), excel_match_array_6_v, ONE).number == 2);

  // When not given a range
  assert(excel_match(EXCEL_NUMBER(10.0), EXCEL_NUMBER(10), EXCEL_NUMBER(0.0)).number == 1);
  assert(excel_match(EXCEL_NUMBER(20.0), EXCEL_NUMBER(10), EXCEL_NUMBER(0.0)).type == ExcelError);
  assert(excel_match(EXCEL_NUMBER(10.0), excel_match_array_1_v, BLANK).number == 1);

  // Test more than on
  // .. numbers
  assert(more_than(ONE,EXCEL_NUMBER(2)).number == false);
  assert(more_than(ONE,ONE).number == false);
  assert(more_than(ONE,EXCEL_NUMBER(0)).number == true);
  // .. booleans
  assert(more_than(FALSE,FALSE).number == false);
  assert(more_than(FALSE,TRUE).number == false);
  assert(more_than(TRUE,FALSE).number == true);
  assert(more_than(TRUE,TRUE).number == false);
  // ..strings
  assert(more_than(EXCEL_STRING("HELLO"),EXCEL_STRING("Ardvark")).number == true);
  assert(more_than(EXCEL_STRING("HELLO"),EXCEL_STRING("world")).number == false);
  assert(more_than(EXCEL_STRING("HELLO"),EXCEL_STRING("hello")).number == false);
  // ..blanks
  assert(more_than(BLANK,ONE).number == false);
  assert(more_than(BLANK,EXCEL_NUMBER(-1)).number == true);
  assert(more_than(ONE,BLANK).number == true);
  assert(more_than(EXCEL_NUMBER(-1),BLANK).number == false);
  // .. of different types
  assert(more_than(TRUE,EXCEL_STRING("Hello")).number == true);
  assert(more_than(FALSE,EXCEL_STRING("Hello")).number == true);
  assert(more_than(EXCEL_STRING("Hello"), ONE).number == true);
  assert(more_than(ONE,EXCEL_STRING("Hello")).number == false);
  assert(more_than(EXCEL_STRING("Hello"), TRUE).number == false);
  assert(more_than(EXCEL_STRING("Hello"), FALSE).number == false);

  // Test less than on
  // .. numbers
  assert(less_than(ONE,EXCEL_NUMBER(2)).number == true);
  assert(less_than(ONE,ONE).number == false);
  assert(less_than(ONE,EXCEL_NUMBER(0)).number == false);
  // .. booleans
  assert(less_than(FALSE,FALSE).number == false);
  assert(less_than(FALSE,TRUE).number == true);
  assert(less_than(TRUE,FALSE).number == false);
  assert(less_than(TRUE,TRUE).number == false);
  // ..strings
  assert(less_than(EXCEL_STRING("HELLO"),EXCEL_STRING("Ardvark")).number == false);
  assert(less_than(EXCEL_STRING("HELLO"),EXCEL_STRING("world")).number == true);
  assert(less_than(EXCEL_STRING("HELLO"),EXCEL_STRING("hello")).number == false);
  // ..blanks
  assert(less_than(BLANK,ONE).number == true);
  assert(less_than(BLANK,EXCEL_NUMBER(-1)).number == false);
  assert(less_than(ONE,BLANK).number == false);
  assert(less_than(EXCEL_NUMBER(-1),BLANK).number == true);
  // .. of different types
  assert(less_than(TRUE,EXCEL_STRING("Hello")).number == false);
  assert(less_than(FALSE,EXCEL_STRING("Hello")).number == false);
  assert(less_than(EXCEL_STRING("Hello"), ONE).number == false);
  assert(less_than(ONE,EXCEL_STRING("Hello")).number == true);
  assert(less_than(EXCEL_STRING("Hello"), TRUE).number == true);
  assert(less_than(EXCEL_STRING("Hello"), FALSE).number == true);

  // Test FIND function
  // ... should find the first occurrence of one string in another, returning :value if the string doesn't match
  assert(find_2(EXCEL_STRING("one"),EXCEL_STRING("onetwothree")).number == 1);
  assert(find_2(EXCEL_STRING("one"),EXCEL_STRING("twoonethree")).number == 4);
  assert(find_2(EXCEL_STRING("one"),EXCEL_STRING("twoonthree")).type == ExcelError);
  // ... should find the first occurrence of one string in another after a given index, returning :value if the string doesn't match
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("onetwothree"),ONE).number == 1);
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("twoonethree"),EXCEL_NUMBER(5)).type == ExcelError);
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_NUMBER(2)).number == 4);
  // ... should be possible for the start_num to be a string, if that string converts to a number
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_STRING("2")).number == 4);
  // ... should return a :value error when given anything but a number as the third argument
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_STRING("a")).type == ExcelError);
  // ... should return a :value error when given a third argument that is less than 1 or greater than the length of the string
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_NUMBER(0)).type == ExcelError);
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_NUMBER(-1)).type == ExcelError);
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("oneone"),EXCEL_NUMBER(7)).type == ExcelError);
  // ... BLANK in the first argument matches any character
  assert(find_2(BLANK,EXCEL_STRING("abcdefg")).number == 1);
  assert(find(BLANK,EXCEL_STRING("abcdefg"),EXCEL_NUMBER(4)).number == 4);
  // ... should treat BLANK in the second argument as an empty string
  assert(find_2(BLANK,BLANK).number == 1);
  assert(find_2(EXCEL_STRING("a"),BLANK).type == ExcelError);
  // ... should return an error if any argument is an error
  assert(find(EXCEL_STRING("one"),EXCEL_STRING("onetwothree"),NA).type == ExcelError);
  assert(find(EXCEL_STRING("one"),NA,ONE).type == ExcelError);
  assert(find(NA,EXCEL_STRING("onetwothree"),ONE).type == ExcelError);

  // Test the IFERROR function
  assert(iferror(EXCEL_STRING("ok"),ONE).type == ExcelString);
  assert(iferror(VALUE,ONE).type == ExcelNumber);

  // Test the ISERR function
  assert(iserr(NA).type == ExcelBoolean);
  assert(iserr(NA).number == 0);
  assert(iserr(DIV0).type == ExcelBoolean);
  assert(iserr(DIV0).type == ExcelBoolean);
  assert(iserr(REF).number == 1);
  assert(iserr(REF).type == ExcelBoolean);
  assert(iserr(VALUE).number == 1);
  assert(iserr(VALUE).type == ExcelBoolean);
  assert(iserr(NAME).number == 1);
  assert(iserr(NAME).number == 1);
  assert(iserr(BLANK).type == ExcelBoolean);
  assert(iserr(BLANK).type == ExcelBoolean);
  assert(iserr(TRUE).type == ExcelBoolean);
  assert(iserr(TRUE).type == ExcelBoolean);
  assert(iserr(FALSE).number == 0);
  assert(iserr(FALSE).number == 0);
  assert(iserr(ONE).number == 0);
  assert(iserr(ONE).number == 0);
  assert(iserr(EXCEL_STRING("Hello")).number == 0);
  assert(iserr(EXCEL_STRING("Hello")).number == 0);

  // Test the ISERROR function
  assert_equal(iserror(NA), TRUE, "ISERROR(NA)");
  assert_equal(iserror(DIV0), TRUE, "ISERROR(DIV0)");
  assert_equal(iserror(REF), TRUE, "ISERROR(REF)");
  assert_equal(iserror(VALUE), TRUE, "ISERROR(VALUE)");
  assert_equal(iserror(NAME), TRUE, "ISERROR(NAME)");
  assert_equal(iserror(BLANK), FALSE, "ISERROR(BLANK)");
  assert_equal(iserror(TRUE), FALSE, "ISERROR(TRUE)");
  assert_equal(iserror(FALSE), FALSE, "ISERROR(FALSE)");
  assert_equal(iserror(ONE), FALSE, "ISERROR(ONE)");
  assert_equal(iserror(EXCEL_STRING("Hello")), FALSE, "ISERROR('Hello')");

  // Test the INDEX function
  ExcelValue index_array_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(20), BLANK };
  ExcelValue index_array_1_v_column = EXCEL_RANGE(index_array_1,3,1);
  ExcelValue index_array_1_v_row = EXCEL_RANGE(index_array_1,1,3);
  ExcelValue index_array_2[] = { BLANK, ONE, EXCEL_NUMBER(10), EXCEL_NUMBER(11), EXCEL_NUMBER(100), EXCEL_NUMBER(101) };
  ExcelValue index_array_2_v = EXCEL_RANGE(index_array_2,3,2);
  // ... if given one argument should return the value at that offset in the range
  assert(excel_index_2(index_array_1_v_column,EXCEL_NUMBER(2.0)).number == 20);
  assert(excel_index_2(index_array_1_v_row,EXCEL_NUMBER(2.0)).number == 20);
  // ... but not if the range is not a single row or single column
  assert(excel_index_2(index_array_2_v,EXCEL_NUMBER(2.0)).type == ExcelError);
  // ... it should return the value in the array at position row_number, column_number
  assert(excel_index(EXCEL_NUMBER(10),ONE,ONE).number == 10);
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(1.0),EXCEL_NUMBER(2.0)).number == 1);
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(2.0),EXCEL_NUMBER(1.0)).number == 10);
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(3.0),EXCEL_NUMBER(1.0)).number == 100);
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(3.0),EXCEL_NUMBER(3.0)).type == ExcelError);
  // ... it should return ZERO not blank, if a blank cell is picked
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(1.0),EXCEL_NUMBER(1.0)).type == ExcelNumber);
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(1.0),EXCEL_NUMBER(1.0)).number == 0);
  assert(excel_index_2(index_array_1_v_row,EXCEL_NUMBER(3.0)).type == ExcelNumber);
  assert(excel_index_2(index_array_1_v_row,EXCEL_NUMBER(3.0)).number == 0);
  // ... it should return the whole row if given a zero column number
  ExcelValue index_result_1_v = excel_index(index_array_2_v,EXCEL_NUMBER(1.0),EXCEL_NUMBER(0.0));
  assert(index_result_1_v.type == ExcelRange);
  assert(index_result_1_v.rows == 1);
  assert(index_result_1_v.columns == 2);
  ExcelValue *index_result_1_a = index_result_1_v.array;
  assert(index_result_1_a[0].number == 0);
  assert(index_result_1_a[1].number == 1);
  // ... it should return the whole column if given a zero row number
  ExcelValue index_result_2_v = excel_index(index_array_2_v,EXCEL_NUMBER(0),EXCEL_NUMBER(1.0));
  assert(index_result_2_v.type == ExcelRange);
  assert(index_result_2_v.rows == 3);
  assert(index_result_2_v.columns == 1);
  ExcelValue *index_result_2_a = index_result_2_v.array;
  assert(index_result_2_a[0].number == 0);
  assert(index_result_2_a[1].number == 10);
  assert(index_result_2_a[2].number == 100);
  // ... it should return a :ref error when given arguments outside array range
  assert(excel_index_2(index_array_1_v_row,EXCEL_NUMBER(-1)).type == ExcelError);
  assert(excel_index_2(index_array_1_v_row,EXCEL_NUMBER(4)).type == ExcelError);
  // ... it should treat BLANK as zero if given as a required row or column number
  assert(excel_index(index_array_2_v,EXCEL_NUMBER(1.0),BLANK).type == ExcelRange);
  assert(excel_index(index_array_2_v,BLANK,EXCEL_NUMBER(2.0)).type == ExcelRange);
  // ... it should return an error if an argument is an error
  assert(excel_index(NA,NA,NA).type == ExcelError);
  // ... it should return a single value if single column and passed zero as column number
  assert(excel_index(index_array_1_v_column,EXCEL_NUMBER(2.0), ZERO).number == 20);
  assert(excel_index(index_array_1_v_row,ZERO, EXCEL_NUMBER(2.0)).number == 20);


  // LEFT(string,[characters])
  // ... should return the left n characters from a string
  assert(strcmp(left_1(EXCEL_STRING("ONE")).string,"O") == 0);
  assert(strcmp(left(EXCEL_STRING("ONE"),ONE).string,"O") == 0);
  assert(strcmp(left(EXCEL_STRING("ONE"),EXCEL_NUMBER(3)).string,"ONE") == 0);
  // ... should turn numbers into strings before processing
  assert(strcmp(left(EXCEL_NUMBER(1.31e12),EXCEL_NUMBER(3)).string, "131") == 0);
  // ... should turn booleans into the words TRUE and FALSE before processing
  assert(strcmp(left(TRUE,EXCEL_NUMBER(3)).string,"TRU") == 0);
  assert(strcmp(left(FALSE,EXCEL_NUMBER(3)).string,"FAL") == 0);
  // ... should return BLANK if given BLANK for either argument
  assert(left(BLANK,EXCEL_NUMBER(3)).type == ExcelEmpty);
  assert(left(EXCEL_STRING("ONE"),BLANK).type == ExcelEmpty);
  // ... should return an error if an argument is an error
  assert(left_1(NA).type == ExcelError);
  assert(left(EXCEL_STRING("ONE"),NA).type == ExcelError);
  assert(left(EXCEL_STRING("ONE"),EXCEL_NUMBER(-10)).type == ExcelError);
  assert_equal(EXCEL_STRING("ONE"), left(EXCEL_STRING("ONE"), EXCEL_NUMBER(100)), "LEFT if number of characters greater than string length");

  // Test less than or equal to
  // .. numbers
  assert(less_than_or_equal(ONE,EXCEL_NUMBER(2)).number == true);
  assert(less_than_or_equal(ONE,ONE).number == true);
  assert(less_than_or_equal(ONE,EXCEL_NUMBER(0)).number == false);
  // .. booleans
  assert(less_than_or_equal(FALSE,FALSE).number == true);
  assert(less_than_or_equal(FALSE,TRUE).number == true);
  assert(less_than_or_equal(TRUE,FALSE).number == false);
  assert(less_than_or_equal(TRUE,TRUE).number == true);
  // ..strings
  assert(less_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("Ardvark")).number == false);
  assert(less_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("world")).number == true);
  assert(less_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("hello")).number == true);
  // ..blanks
  assert(less_than_or_equal(BLANK,ONE).number == true);
  assert(less_than_or_equal(BLANK,EXCEL_NUMBER(-1)).number == false);
  assert(less_than_or_equal(ONE,BLANK).number == false);
  assert(less_than_or_equal(EXCEL_NUMBER(-1),BLANK).number == true);

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
  assert(mod(EXCEL_NUMBER(10), EXCEL_NUMBER(3)).number == 1.0);
  assert(mod(EXCEL_NUMBER(10), EXCEL_NUMBER(5)).number == 0.0);
  // ... should be possible for the the arguments to be strings, if they convert to a number
  assert(mod(EXCEL_STRING("3.5"),EXCEL_STRING("2")).number == 1.5);
  // ... should treat BLANK as zero
  assert(mod(BLANK,EXCEL_NUMBER(10)).number == 0);
  assert(mod(EXCEL_NUMBER(10),BLANK).type == ExcelError);
  assert(mod(BLANK,BLANK).type == ExcelError);
  // ... should treat true as 1 and FALSE as 0
  assert((mod(EXCEL_NUMBER(1.1),TRUE).number - 0.1) < 0.001);
  assert(mod(EXCEL_NUMBER(1.1),FALSE).type == ExcelError);
  assert(mod(FALSE,EXCEL_NUMBER(10)).number == 0);
  // ... should return an error when given inappropriate arguments
  assert(mod(EXCEL_STRING("Asdasddf"),EXCEL_STRING("adsfads")).type == ExcelError);
  // ... should return an error if an argument is an error
  assert(mod(EXCEL_NUMBER(1),VALUE).type == ExcelError);
  assert(mod(VALUE,EXCEL_NUMBER(1)).type == ExcelError);
  assert(mod(VALUE,VALUE).type == ExcelError);

  // Test more than or equal to on
  // .. numbers
  assert(more_than_or_equal(ONE,EXCEL_NUMBER(2)).number == false);
  assert(more_than_or_equal(ONE,ONE).number == true);
  assert(more_than_or_equal(ONE,EXCEL_NUMBER(0)).number == true);
  // .. booleans
  assert(more_than_or_equal(FALSE,FALSE).number == true);
  assert(more_than_or_equal(FALSE,TRUE).number == false);
  assert(more_than_or_equal(TRUE,FALSE).number == true);
  assert(more_than_or_equal(TRUE,TRUE).number == true);
  // ..strings
  assert(more_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("Ardvark")).number == true);
  assert(more_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("world")).number == false);
  assert(more_than_or_equal(EXCEL_STRING("HELLO"),EXCEL_STRING("hello")).number == true);
  // ..blanks
  assert(more_than_or_equal(BLANK,BLANK).number == true);
  assert(more_than_or_equal(BLANK,ONE).number == false);
  assert(more_than_or_equal(BLANK,EXCEL_NUMBER(-1)).number == true);
  assert(more_than_or_equal(ONE,BLANK).number == true);
  assert(more_than_or_equal(EXCEL_NUMBER(-1),BLANK).number == false);

  // Test negative
  // ... should return the negative of its arguments
  assert(negative(EXCEL_NUMBER(1)).number == -1);
  assert(negative(EXCEL_NUMBER(-1)).number == 1);
  // ... should treat strings that only contain numbers as numbers
  assert(negative(EXCEL_STRING("10")).number == -10);
  assert(negative(EXCEL_STRING("-1.3")).number == 1.3);
  // ... should return an error when given inappropriate arguments
  assert(negative(EXCEL_STRING("Asdasddf")).type == ExcelError);
  // ... should treat BLANK as zero
  assert(negative(BLANK).number == 0);

  // Test PMT(rate,number_of_periods,present_value) - optional arguments not yet implemented
  // ... should calculate the monthly payment required for a given principal, interest rate and loan period
  assert((pmt(EXCEL_NUMBER(0.1),EXCEL_NUMBER(10),EXCEL_NUMBER(100)).number - -16.27) < 0.01);
  assert((pmt(EXCEL_NUMBER(0.0123),EXCEL_NUMBER(99.1),EXCEL_NUMBER(123.32)).number - -2.159) < 0.01);
  assert((pmt(EXCEL_NUMBER(0),EXCEL_NUMBER(2),EXCEL_NUMBER(10)).number - -5) < 0.01);
  assert((pmt_4(EXCEL_NUMBER(0),EXCEL_NUMBER(2),EXCEL_NUMBER(10), EXCEL_NUMBER(0)).number - -5) < 0.01);
  assert((pmt_5(EXCEL_NUMBER(0),EXCEL_NUMBER(2),EXCEL_NUMBER(10), EXCEL_NUMBER(0), EXCEL_NUMBER(0)).number - -5) < 0.01);

  // Test power
  // ... should return power of its arguments
  assert(power(EXCEL_NUMBER(2),EXCEL_NUMBER(3)).number == 8);
  assert(power(EXCEL_NUMBER(4.0),EXCEL_NUMBER(0.5)).number == 2.0);
  assert(power(EXCEL_NUMBER(-4.0),EXCEL_NUMBER(0.5)).type == ExcelError);

  // Test round
  assert(excel_round(EXCEL_NUMBER(1.1), EXCEL_NUMBER(0)).number == 1.0);
  assert(excel_round(EXCEL_NUMBER(1.5), EXCEL_NUMBER(0)).number == 2.0);
  assert(excel_round(EXCEL_NUMBER(1.56),EXCEL_NUMBER(1)).number == 1.6);
  assert(excel_round(EXCEL_NUMBER(-1.56),EXCEL_NUMBER(1)).number == -1.6);

  // Test rounddown
  assert(rounddown(EXCEL_NUMBER(1.1), EXCEL_NUMBER(0)).number == 1.0);
  assert(rounddown(EXCEL_NUMBER(1.5), EXCEL_NUMBER(0)).number == 1.0);
  assert(rounddown(EXCEL_NUMBER(1.56),EXCEL_NUMBER(1)).number == 1.5);
  assert(rounddown(EXCEL_NUMBER(-1.56),EXCEL_NUMBER(1)).number == -1.5);

  // Test int
  assert(excel_int(EXCEL_NUMBER(8.9)).number == 8.0);
  assert(excel_int(EXCEL_NUMBER(-8.9)).number == -9.0);

  // Test roundup
  assert(roundup(EXCEL_NUMBER(1.1), EXCEL_NUMBER(0)).number == 2.0);
  assert(roundup(EXCEL_NUMBER(1.5), EXCEL_NUMBER(0)).number == 2.0);
  assert(roundup(EXCEL_NUMBER(1.56),EXCEL_NUMBER(1)).number == 1.6);
  assert(roundup(EXCEL_NUMBER(-1.56),EXCEL_NUMBER(1)).number == -1.6);

  // Test string joining
  ExcelValue string_join_array_1[] = {EXCEL_STRING("Hello "), EXCEL_STRING("world")};
  ExcelValue string_join_array_2[] = {EXCEL_STRING("Hello "), EXCEL_STRING("world"), EXCEL_STRING("!")};
  ExcelValue string_join_array_3[] = {EXCEL_STRING("Top "), EXCEL_NUMBER(10.0)};
  ExcelValue string_join_array_4[] = {EXCEL_STRING("Top "), EXCEL_NUMBER(10.5)};
  ExcelValue string_join_array_5[] = {EXCEL_STRING("Top "), TRUE, FALSE};
  // ... should return a string by combining its arguments
  // inspect_excel_value(string_join(2, string_join_array_1));
  assert(string_join(2, string_join_array_1).string[6] == 'w');
  assert(string_join(2, string_join_array_1).string[11] == '\0');
  // ... should cope with an arbitrary number of arguments
  assert(string_join(3, string_join_array_2).string[11] == '!');
  assert_equal(EXCEL_STRING("Top 10"), string_join(2, string_join_array_3), "String join with numbers");
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
  ExcelValue string_join_array_6[] = {EXCEL_STRING("0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"), EXCEL_STRING("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789")};
  assert(string_join(2, string_join_array_6).string[0] == '0');
  // Even ones that are greater than 200 characters
  ExcelValue string_join_array_6b[] = {EXCEL_STRING("Can we increase crop yields? Between 1987 and 2007 they increased by around 1.9% per year globally, but over recent years the        annual % increase has been lower than this. Current average food energy yields are "), EXCEL_STRING("Can we increase crop yields? Between 1987 and 2007 they increased by around 1.9% per year globally, but over recent years the        annual % increase has been lower than this. Current average food energy yields are ")};
  assert_equal(EXCEL_STRING("Can we increase crop yields? Between 1987 and 2007 they increased by around 1.9% per year globally, but over recent years the        annual % increase has been lower than this. Current average food energy yields are Can we increase crop yields? Between 1987 and 2007 they increased by around 1.9% per year globally, but over recent years the        annual % increase has been lower than this. Current average food energy yields are "), string_join(2, string_join_array_6b), "String join > 200 character");
  // Should deal with some edge cases
  ExcelValue string_join_array_7[] = { NA };
  assert_equal(NA, string_join(1, string_join_array_7), "String_join should return an error when passed an error");
  ExcelValue string_join_array_8[] = { EXCEL_RANGE(string_join_array_7, 1, 1) };
  assert_equal(VALUE, string_join(1, string_join_array_8), "String_join should return VALUE when passed a range");

  // Test SUBTOTAL function
  ExcelValue subtotal_array_1[] = {EXCEL_NUMBER(10),EXCEL_NUMBER(100),BLANK};
  ExcelValue subtotal_array_1_v = EXCEL_RANGE(subtotal_array_1,3,1);
  ExcelValue subtotal_array_2[] = {EXCEL_NUMBER(1),EXCEL_STRING("two"),subtotal_array_1_v};

  // EXCEL_NUMBER(1.0);
  // inspect_excel_value(EXCEL_NUMBER(1.0));
  // inspect_excel_value(EXCEL_RANGE(subtotal_array_2,3,1));
  // inspect_excel_value(subtotal(EXCEL_NUMBER(1.0),3,subtotal_array_2));

  assert(subtotal(EXCEL_NUMBER(1.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(EXCEL_NUMBER(2.0),3,subtotal_array_2).number == 3);
  assert(subtotal(EXCEL_NUMBER(3.0),7, count_a_test_array_1).number == 6);
  assert(subtotal(EXCEL_NUMBER(3.0),3,subtotal_array_2).number == 4);
  assert(subtotal(EXCEL_NUMBER(9.0),3,subtotal_array_2).number == 111);
  assert(subtotal(EXCEL_NUMBER(101.0),3,subtotal_array_2).number == 111.0/3.0);
  assert(subtotal(EXCEL_NUMBER(102.0),3,subtotal_array_2).number == 3);
  assert(subtotal(EXCEL_NUMBER(103.0),3,subtotal_array_2).number == 4);
  assert(subtotal(EXCEL_NUMBER(109.0),3,subtotal_array_2).number == 111);

  // Test SUMIFS function
  ExcelValue sumifs_array_1[] = {EXCEL_NUMBER(10),EXCEL_NUMBER(100),BLANK};
  ExcelValue sumifs_array_1_v = EXCEL_RANGE(sumifs_array_1,3,1);
  ExcelValue sumifs_array_2[] = {EXCEL_STRING("pear"),EXCEL_STRING("bear"),EXCEL_STRING("apple")};
  ExcelValue sumifs_array_2_v = EXCEL_RANGE(sumifs_array_2,3,1);
  ExcelValue sumifs_array_3[] = {EXCEL_NUMBER(1),EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(4),EXCEL_NUMBER(5),EXCEL_NUMBER(5)};
  ExcelValue sumifs_array_3_v = EXCEL_RANGE(sumifs_array_3,6,1);
  ExcelValue sumifs_array_4[] = {EXCEL_STRING("CO2"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CO2")};
  ExcelValue sumifs_array_4_v = EXCEL_RANGE(sumifs_array_4,6,1);
  ExcelValue sumifs_array_5[] = {EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_NUMBER(4),EXCEL_NUMBER(4),EXCEL_NUMBER(5)};
  ExcelValue sumifs_array_5_v = EXCEL_RANGE(sumifs_array_5,6,1);

  // ... should only sum values that meet all of the criteria
  ExcelValue sumifs_array_6[] = { sumifs_array_1_v, EXCEL_NUMBER(10), sumifs_array_2_v, EXCEL_STRING("Bear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_6).number == 0.0);

  ExcelValue sumifs_array_7[] = { sumifs_array_1_v, EXCEL_NUMBER(10), sumifs_array_2_v, EXCEL_STRING("Pear") };
  assert(sumifs(sumifs_array_1_v,4,sumifs_array_7).number == 10.0);

  // ... should work when single cells are given where ranges expected
  ExcelValue sumifs_array_8[] = { EXCEL_STRING("CAR"), EXCEL_STRING("CAR"), EXCEL_STRING("FCV"), EXCEL_STRING("FCV")};
  assert(sumifs(EXCEL_NUMBER(0.143897265452564), 4, sumifs_array_8).number == 0.143897265452564);

  // ... should match numbers with strings that contain numbers
  ExcelValue sumifs_array_9[] = { EXCEL_NUMBER(10), EXCEL_STRING("10.0")};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_9).number == 100);

  ExcelValue sumifs_array_9b[] = { EXCEL_STRING("10"), EXCEL_NUMBER(10.0)};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_9b).number == 100);

  ExcelValue sumifs_array_10[] = { sumifs_array_4_v, EXCEL_STRING("CO2"), sumifs_array_5_v, EXCEL_NUMBER(2)};
  assert(sumifs(sumifs_array_3_v,4, sumifs_array_10).number == 0);

  // ... should match with strings that contain criteria
  ExcelValue sumifs_array_10a[] = { sumifs_array_3_v, EXCEL_STRING("=5")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10a).number == 10);

  ExcelValue sumifs_array_10b[] = { sumifs_array_3_v, EXCEL_STRING("<>3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10b).number == 17);

  ExcelValue sumifs_array_10c[] = { sumifs_array_3_v, EXCEL_STRING("<3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10c).number == 3);

  ExcelValue sumifs_array_10d[] = { sumifs_array_3_v, EXCEL_STRING("<=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10d).number == 6);

  ExcelValue sumifs_array_10e[] = { sumifs_array_3_v, EXCEL_STRING(">3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10e).number == 14);

  ExcelValue sumifs_array_10f[] = { sumifs_array_3_v, EXCEL_STRING(">=3")};
  assert(sumifs(sumifs_array_3_v,2, sumifs_array_10f).number == 17);

  // ... BLANK in check range should match empty strings, BLANK in criteria should match zero
  ExcelValue sumifs_array_11[] = { BLANK, EXCEL_NUMBER(0)};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_11).number == 0);

  ExcelValue sumifs_array_11b[] = { BLANK, EXCEL_STRING("")};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_11b).number == 100);

  ExcelValue sumifs_array_11c[] = { EXCEL_STRING(""), BLANK};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_11c).number == 0);

  ExcelValue sumifs_array_12[] = {EXCEL_NUMBER(0), BLANK};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_12).number == 100);

  ExcelValue sumifs_array_13[] = {BLANK, BLANK};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_13).number == 0);

  ExcelValue sumifs_array_14[] = {EXCEL_NUMBER(10), BLANK};
  assert(sumifs(EXCEL_NUMBER(100),2,sumifs_array_14).number == 0);

  // ... should return an error if range argument is an error
  ExcelValue sumifs_array_15[] = {ONE, ONE};
  assert(sumifs(REF,2,sumifs_array_15).type == ExcelError);

  // Test COUNTIFS function
  ExcelValue countifs_array_1[] = {EXCEL_NUMBER(10),EXCEL_NUMBER(100),BLANK};
  ExcelValue countifs_array_1_v = EXCEL_RANGE(countifs_array_1,3,1);
  ExcelValue countifs_array_2[] = {EXCEL_STRING("pear"),EXCEL_STRING("bear"),EXCEL_STRING("apple")};
  ExcelValue countifs_array_2_v = EXCEL_RANGE(countifs_array_2,3,1);
  ExcelValue countifs_array_3[] = {EXCEL_NUMBER(1),EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(4),EXCEL_NUMBER(5),EXCEL_NUMBER(5)};
  ExcelValue countifs_array_3_v = EXCEL_RANGE(countifs_array_3,6,1);
  ExcelValue countifs_array_4[] = {EXCEL_STRING("CO2"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CO2")};
  ExcelValue countifs_array_4_v = EXCEL_RANGE(countifs_array_4,6,1);
  ExcelValue countifs_array_5[] = {EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_NUMBER(4),EXCEL_NUMBER(4),EXCEL_NUMBER(5)};
  ExcelValue countifs_array_5_v = EXCEL_RANGE(countifs_array_5,6,1);

  // ... should only sum values that meet all of the criteria
  ExcelValue countifs_array_6[] = { countifs_array_1_v, EXCEL_NUMBER(10), countifs_array_2_v, EXCEL_STRING("Bear") };
  assert(countifs(4,countifs_array_6).number == 0.0);

  ExcelValue countifs_array_7[] = { countifs_array_1_v, EXCEL_NUMBER(10), countifs_array_2_v, EXCEL_STRING("Pear") };
  assert(countifs(4,countifs_array_7).number == 1.0);

  // ... should work when single cells are given where ranges expected
  ExcelValue countifs_array_8[] = { EXCEL_STRING("CAR"), EXCEL_STRING("CAR"), EXCEL_STRING("FCV"), EXCEL_STRING("FCV")};
  assert(countifs(4, countifs_array_8).number == 1.0);

  // ... should match numbers with strings that contain numbers
  ExcelValue countifs_array_9[] = { EXCEL_NUMBER(10), EXCEL_STRING("10.0")};
  assert(countifs(2,countifs_array_9).number == 1.0);

  ExcelValue countifs_array_9b[] = { EXCEL_STRING("10"), EXCEL_NUMBER(10.0)};
  assert(countifs(2,countifs_array_9b).number == 1.0);

  ExcelValue countifs_array_10[] = { countifs_array_4_v, EXCEL_STRING("CO2"), countifs_array_5_v, EXCEL_NUMBER(2)};
  assert(countifs(4, countifs_array_10).number == 0.0);

  // ... should match with strings that contain criteria
  ExcelValue countifs_array_10a[] = { countifs_array_3_v, EXCEL_STRING("=5")};
  assert(countifs(2, countifs_array_10a).number == 2.0);

  ExcelValue countifs_array_10b[] = { countifs_array_3_v, EXCEL_STRING("<>3")};
  assert(countifs(2, countifs_array_10b).number == 5.0);

  ExcelValue countifs_array_10c[] = { countifs_array_3_v, EXCEL_STRING("<3")};
  assert(countifs(2, countifs_array_10c).number == 2.0);

  ExcelValue countifs_array_10d[] = { countifs_array_3_v, EXCEL_STRING("<=3")};
  assert(countifs(2, countifs_array_10d).number == 3.0);

  ExcelValue countifs_array_10e[] = { countifs_array_3_v, EXCEL_STRING(">3")};
  assert(countifs(2, countifs_array_10e).number == 3.0);

  ExcelValue countifs_array_10f[] = { countifs_array_3_v, EXCEL_STRING(">=3")};
  assert(countifs(2, countifs_array_10f).number == 4.0);

  // ... BLANK in check range should match empty strings, BLANK in criteria should match zero
  ExcelValue countifs_array_11[] = { BLANK, EXCEL_NUMBER(0)};
  assert(countifs(2,countifs_array_11).number == 0);

  ExcelValue countifs_array_11b[] = { BLANK, EXCEL_STRING("")};
  assert(countifs(2,countifs_array_11b).number == 1);

  ExcelValue countifs_array_11c[] = { EXCEL_STRING(""), BLANK};
  assert(countifs(2,countifs_array_11c).number == 0);

  ExcelValue countifs_array_12[] = {EXCEL_NUMBER(0), BLANK};
  assert(countifs(2,countifs_array_12).number == 1);

  ExcelValue countifs_array_13[] = {BLANK, BLANK};
  assert(countifs(2,countifs_array_13).number == 0);

  ExcelValue countifs_array_14[] = {EXCEL_NUMBER(10), BLANK};
  assert(countifs(2,countifs_array_14).number == 0);

  // Test SUMIF
  // ... where there is only a check range
  assert(sumif_2(sumifs_array_1_v,EXCEL_STRING(">0")).number == 110.0);
  assert(sumif_2(sumifs_array_1_v,EXCEL_STRING(">10")).number == 100.0);
  assert(sumif_2(sumifs_array_1_v,EXCEL_STRING("<100")).number == 10.0);

  // ... where there is a seprate sum range
  ExcelValue sumif_array_1[] = {EXCEL_NUMBER(15),EXCEL_NUMBER(20), EXCEL_NUMBER(30)};
  ExcelValue sumif_array_1_v = EXCEL_RANGE(sumif_array_1,3,1);
  assert(sumif(sumifs_array_1_v,EXCEL_STRING("10"),sumif_array_1_v).number == 15);


  // Test SUMPRODUCT
  ExcelValue sumproduct_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), BLANK};
  ExcelValue sumproduct_2[] = { BLANK, EXCEL_NUMBER(100), EXCEL_NUMBER(10), BLANK};
  ExcelValue sumproduct_3[] = { BLANK };
  ExcelValue sumproduct_4[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), EXCEL_NUMBER(1000)};
  ExcelValue sumproduct_5[] = { EXCEL_NUMBER(1), EXCEL_NUMBER(2), EXCEL_NUMBER(3)};
  ExcelValue sumproduct_6[] = { EXCEL_NUMBER(1), EXCEL_NUMBER(2), EXCEL_NUMBER(4), EXCEL_NUMBER(5)};
  ExcelValue sumproduct_7[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(20), EXCEL_NUMBER(40), EXCEL_NUMBER(50)};
  ExcelValue sumproduct_8[] = { EXCEL_NUMBER(11), EXCEL_NUMBER(21), EXCEL_NUMBER(41), EXCEL_NUMBER(51)};
  ExcelValue sumproduct_9[] = { BLANK, BLANK };

  ExcelValue sumproduct_1_v = EXCEL_RANGE( sumproduct_1, 3, 1);
  ExcelValue sumproduct_2_v = EXCEL_RANGE( sumproduct_2, 3, 1);
  ExcelValue sumproduct_3_v = EXCEL_RANGE( sumproduct_3, 1, 1);
  // ExcelValue sumproduct_4_v = EXCEL_RANGE( sumproduct_4, 1, 3); // Unused
  ExcelValue sumproduct_5_v = EXCEL_RANGE( sumproduct_5, 3, 1);
  ExcelValue sumproduct_6_v = EXCEL_RANGE( sumproduct_6, 2, 2);
  ExcelValue sumproduct_7_v = EXCEL_RANGE( sumproduct_7, 2, 2);
  ExcelValue sumproduct_8_v = EXCEL_RANGE( sumproduct_8, 2, 2);
  ExcelValue sumproduct_9_v = EXCEL_RANGE( sumproduct_9, 2, 1);

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
  ExcelValue sumproducta_6[] = {BLANK,EXCEL_NUMBER(1)};
  assert(sumproduct(2,sumproducta_6).type == ExcelError);

  // ... should ignore non-numeric values within an array
  ExcelValue sumproducta_7[] = {sumproduct_9_v, sumproduct_9_v};
  assert(sumproduct(2,sumproducta_7).number == 0);

  // ... should return an error if an argument is an error
  ExcelValue sumproducta_8[] = {VALUE};
  assert(sumproduct(1,sumproducta_8).type == ExcelError);

  // Test PRODUCT
  ExcelValue product_1[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), BLANK};
  ExcelValue product_2[] = { BLANK, EXCEL_NUMBER(100), EXCEL_NUMBER(10), BLANK};
  ExcelValue product_3[] = { BLANK };
  ExcelValue product_4[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(100), EXCEL_NUMBER(1000)};
  ExcelValue product_5[] = { EXCEL_NUMBER(1), EXCEL_NUMBER(2), EXCEL_NUMBER(3)};
  ExcelValue product_6[] = { EXCEL_NUMBER(1), EXCEL_NUMBER(2), EXCEL_NUMBER(4), EXCEL_NUMBER(5)};
  ExcelValue product_7[] = { EXCEL_NUMBER(10), EXCEL_NUMBER(20), EXCEL_NUMBER(40), EXCEL_NUMBER(50)};
  ExcelValue product_8[] = { EXCEL_NUMBER(11), EXCEL_NUMBER(21), EXCEL_NUMBER(41), EXCEL_NUMBER(51)};
  ExcelValue product_9[] = { BLANK, BLANK };

  ExcelValue product_1_v = EXCEL_RANGE( product_1, 3, 1);
  ExcelValue product_2_v = EXCEL_RANGE( product_2, 3, 1);
  ExcelValue product_3_v = EXCEL_RANGE( product_3, 1, 1);
  // ExcelValue product_4_v = EXCEL_RANGE( product_4, 1, 3); // Unused
  ExcelValue product_5_v = EXCEL_RANGE( product_5, 3, 1);
  ExcelValue product_6_v = EXCEL_RANGE( product_6, 2, 2);
  ExcelValue product_7_v = EXCEL_RANGE( product_7, 2, 2);
  ExcelValue product_8_v = EXCEL_RANGE( product_8, 2, 2);
  ExcelValue product_9_v = EXCEL_RANGE( product_9, 2, 1);

  // ... should multiply together the elements in row or column areas given as arguments
  ExcelValue producta_1[] = {product_1_v, product_2_v};
  assert(product(2,producta_1).number == 10*100*100*10);

  // ... should work when miss-matched array sizes
  ExcelValue producta_2[] = {product_1_v, product_3_v};
  assert(product(2,producta_2).number == 10 * 100);

  // ... if all its arguments are single values, should multiply them together
  // ExcelValue *producta_3 = product_4;
  assert(product(3,product_4).number == 10*100*1000);

  // ... if it only has one range as an argument, should multiply its elements together
  ExcelValue producta_4[] = {product_5_v};
  assert(product(1,producta_4).number == 1 * 2 * 3);

  // ... if given multi row and column areas as arguments, should multipy the corresponding cell in each area
  // NB: Repeating this test from SUMPRODUCT, doesn't matter really with multiplication
  ExcelValue producta_5[] = {product_6_v, product_7_v, product_8_v};
  // NB: The 1.0 at the start is important, otherwise RHS will be an int with does not equal the double
  assert(product(3,producta_5).number == (1.0*2*4*5)*(10*20*40*50)*(11*21*41*51));

  // ... should ignore BLANK values outside of an array
  ExcelValue producta_6[] = {BLANK,EXCEL_NUMBER(1)};
  assert(product(2,producta_6).type == 1);

  // ... should ignore non-numeric values within an array
  ExcelValue producta_7[] = {product_9_v, product_9_v};
  assert(product(2,producta_7).number == 0);

  // ... should return an error if an argument is an error
  ExcelValue producta_8[] = {VALUE};
  assert(product(1,producta_8).type == ExcelError);

  // Test VLOOKUP
  ExcelValue vlookup_a1[] = {EXCEL_NUMBER(1),EXCEL_NUMBER(10),EXCEL_NUMBER(2),EXCEL_NUMBER(20),EXCEL_NUMBER(3),EXCEL_NUMBER(30)};
  ExcelValue vlookup_a2[] = {EXCEL_STRING("hello"),EXCEL_NUMBER(10),EXCEL_NUMBER(2),EXCEL_NUMBER(20),EXCEL_NUMBER(3),EXCEL_NUMBER(30)};
  ExcelValue vlookup_a3[] = {BLANK,EXCEL_NUMBER(10),EXCEL_NUMBER(2),EXCEL_NUMBER(20),EXCEL_NUMBER(3),EXCEL_NUMBER(30)};
  ExcelValue vlookup_a1_v = EXCEL_RANGE(vlookup_a1,3,2);
  ExcelValue vlookup_a2_v = EXCEL_RANGE(vlookup_a2,3,2);
  ExcelValue vlookup_a3_v = EXCEL_RANGE(vlookup_a3,3,2);
  // ... should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument
  assert(vlookup_3(EXCEL_NUMBER(2.0),vlookup_a1_v,EXCEL_NUMBER(2)).number == 20);
  assert(vlookup_3(EXCEL_NUMBER(1.5),vlookup_a1_v,EXCEL_NUMBER(2)).number == 10);
  assert(vlookup_3(EXCEL_NUMBER(0.5),vlookup_a1_v,EXCEL_NUMBER(2)).type == ExcelError);
  assert(vlookup_3(EXCEL_NUMBER(10),vlookup_a1_v,EXCEL_NUMBER(2)).number == 30);
  assert(vlookup_3(EXCEL_NUMBER(2.6),vlookup_a1_v,EXCEL_NUMBER(2)).number == 20);
  // ... has a four argument variant that matches the lookup type
  assert(vlookup(EXCEL_NUMBER(2.6),vlookup_a1_v,EXCEL_NUMBER(2),TRUE).number == 20);
  assert(vlookup(EXCEL_NUMBER(2.6),vlookup_a1_v,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(vlookup(EXCEL_STRING("HELLO"),vlookup_a2_v,EXCEL_NUMBER(2),FALSE).number == 10);
  assert(vlookup(EXCEL_STRING("HELMP"),vlookup_a2_v,EXCEL_NUMBER(2),TRUE).number == 10);
  // .. the four argument variant should accept 0 and 1 instead of TRUE and FALSE
  assert(vlookup(EXCEL_STRING("HELLO"),vlookup_a2_v,EXCEL_NUMBER(2),ZERO).number == 10);
  assert(vlookup(EXCEL_STRING("HELMP"),vlookup_a2_v,EXCEL_NUMBER(2),ONE).number == 10);
  // ... BLANK should not match with anything" do
  assert(vlookup_3(BLANK,vlookup_a3_v,EXCEL_NUMBER(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(vlookup(VALUE,vlookup_a1_v,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(vlookup(EXCEL_NUMBER(2.0),VALUE,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(vlookup(EXCEL_NUMBER(2.0),vlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(vlookup(EXCEL_NUMBER(2.0),vlookup_a1_v,EXCEL_NUMBER(2),VALUE).type == ExcelError);
  assert(vlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);

  // Test HLOOKUP
  ExcelValue hlookup_a1[] = {EXCEL_NUMBER(1),EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(10),EXCEL_NUMBER(20),EXCEL_NUMBER(30)};
  ExcelValue hlookup_a2[] = {EXCEL_STRING("hello"),EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(10),EXCEL_NUMBER(20),EXCEL_NUMBER(30)};
  ExcelValue hlookup_a3[] = {BLANK,EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(10),EXCEL_NUMBER(20),EXCEL_NUMBER(30)};
  ExcelValue hlookup_a1_v = EXCEL_RANGE(hlookup_a1,2,3);
  ExcelValue hlookup_a2_v = EXCEL_RANGE(hlookup_a2,2,3);
  ExcelValue hlookup_a3_v = EXCEL_RANGE(hlookup_a3,2,3);
  // ... should match the first argument against the first column of the table in the second argument, returning the value in the column specified by the third argument
  assert(hlookup_3(EXCEL_NUMBER(2.0),hlookup_a1_v,EXCEL_NUMBER(2)).number == 20);
  assert(hlookup_3(EXCEL_NUMBER(1.5),hlookup_a1_v,EXCEL_NUMBER(2)).number == 10);
  assert(hlookup_3(EXCEL_NUMBER(0.5),hlookup_a1_v,EXCEL_NUMBER(2)).type == ExcelError);
  assert(hlookup_3(EXCEL_NUMBER(10),hlookup_a1_v,EXCEL_NUMBER(2)).number == 30);
  assert(hlookup_3(EXCEL_NUMBER(2.6),hlookup_a1_v,EXCEL_NUMBER(2)).number == 20);
  // ... has a four argument variant that matches the lookup type
  assert(hlookup(EXCEL_NUMBER(2.6),hlookup_a1_v,EXCEL_NUMBER(2),TRUE).number == 20);
  assert(hlookup(EXCEL_NUMBER(2.6),hlookup_a1_v,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(hlookup(EXCEL_STRING("HELLO"),hlookup_a2_v,EXCEL_NUMBER(2),FALSE).number == 10);
  assert(hlookup(EXCEL_STRING("HELMP"),hlookup_a2_v,EXCEL_NUMBER(2),TRUE).number == 10);
  // ... that four argument variant should accept 0 or 1 for the lookup type
  assert(hlookup(EXCEL_NUMBER(2.6),hlookup_a1_v,EXCEL_NUMBER(2),ONE).number == 20);
  assert(hlookup(EXCEL_NUMBER(2.6),hlookup_a1_v,EXCEL_NUMBER(2),ZERO).type == ExcelError);
  assert(hlookup(EXCEL_STRING("HELLO"),hlookup_a2_v,EXCEL_NUMBER(2),ZERO).number == 10);
  assert(hlookup(EXCEL_STRING("HELMP"),hlookup_a2_v,EXCEL_NUMBER(2),ONE).number == 10);
  // ... BLANK should not match with anything" do
  assert(hlookup_3(BLANK,hlookup_a3_v,EXCEL_NUMBER(2)).type == ExcelError);
  // ... should return an error if an argument is an error" do
  assert(hlookup(VALUE,hlookup_a1_v,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(hlookup(EXCEL_NUMBER(2.0),VALUE,EXCEL_NUMBER(2),FALSE).type == ExcelError);
  assert(hlookup(EXCEL_NUMBER(2.0),hlookup_a1_v,VALUE,FALSE).type == ExcelError);
  assert(hlookup(EXCEL_NUMBER(2.0),hlookup_a1_v,EXCEL_NUMBER(2),VALUE).type == ExcelError);
  assert(hlookup(VALUE,VALUE,VALUE,VALUE).type == ExcelError);

  // Test SUM
  ExcelValue sum_array_0[] = {EXCEL_NUMBER(1084.4557258064517),EXCEL_NUMBER(32.0516914516129),EXCEL_NUMBER(137.36439193548387)};
  ExcelValue sum_array_0_v = EXCEL_RANGE(sum_array_0,3,1);
  ExcelValue sum_array_1[] = {sum_array_0_v};
  assert(sum(1,sum_array_1).number == 1253.8718091935484);

  // Test PV
  assert((int) pv_3(EXCEL_NUMBER(0.03), EXCEL_NUMBER(12), EXCEL_NUMBER(100)).number == -995);
  assert((int) pv_4(EXCEL_NUMBER(0.03), EXCEL_NUMBER(12), EXCEL_NUMBER(-100), EXCEL_NUMBER(100)).number == 925);
  assert((int) pv_5(EXCEL_NUMBER(0.03), EXCEL_NUMBER(12), EXCEL_NUMBER(-100), EXCEL_NUMBER(-100), EXCEL_NUMBER(1)).number == 1095);

  // Test TEXT
  assert(strcmp(text(EXCEL_NUMBER(1.0), EXCEL_STRING("0%")).string, "100%") == 0);
  assert(strcmp(text(EXCEL_STRING("1"), EXCEL_STRING("0%")).string, "100%") == 0);
  assert(strcmp(text(EXCEL_STRING("0.00251"), EXCEL_STRING("0.0%")).string, "0.3%") == 0);
  assert(strcmp(text(BLANK, EXCEL_STRING("0%")).string, "0%") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.0), BLANK).string, "") == 0);
  assert(strcmp(text(EXCEL_STRING("ASGASD"), EXCEL_STRING("0%")).string, "ASGASD") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.1518), EXCEL_STRING("0")).string, "1") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.1518), ZERO).string, "1") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.1518), EXCEL_STRING("0.0")).string, "1.2") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.1518), EXCEL_STRING("0.00")).string, "1.15") == 0);
  assert(strcmp(text(EXCEL_NUMBER(1.1518), EXCEL_STRING("0.000")).string, "1.152") == 0);
  assert(strcmp(text(EXCEL_NUMBER(12.51), EXCEL_STRING("0000")).string, "0013") == 0);
  assert(strcmp(text(EXCEL_NUMBER(125101), EXCEL_STRING("0000")).string, "125101") == 0);
  assert(strcmp(text(EXCEL_NUMBER(123456789.123456), EXCEL_STRING("#,##")).string, "123,456,789") == 0);
  assert(strcmp(text(EXCEL_NUMBER(123456789.123456), EXCEL_STRING("#,##0")).string, "123,456,789") == 0);
  assert(strcmp(text(EXCEL_NUMBER(123456789.123456), EXCEL_STRING("#,##0.0")).string, "123,456,789.1") == 0);
  assert(strcmp(text(EXCEL_NUMBER(123456789.123456), EXCEL_STRING("!#,##0.0")).string, "Text format not recognised") == 0);

  // Test LOG
  // One argument variant assumes LOG base 10
  assert(excel_log(EXCEL_NUMBER(10)).number == 1);
  assert(excel_log(EXCEL_NUMBER(100)).number == 2);
  assert(excel_log(EXCEL_NUMBER(0)).type == ExcelError);
  // Two argument variant allows LOG base to be specified
  assert(excel_log_2(EXCEL_NUMBER(8),EXCEL_NUMBER(2)).number == 3.0);
  assert(excel_log_2(EXCEL_NUMBER(8),EXCEL_NUMBER(0)).type == ExcelError);

  // Test LN
  assert(ln(EXCEL_NUMBER(10)).number == 2.302585092994046);
  assert(ln(EXCEL_NUMBER(8)).number == 2.0794415416798357);
  assert(ln(EXCEL_NUMBER(0)).type == ExcelError);
  assert(ln(EXCEL_NUMBER(-1)).type == ExcelError);

  // Test excel_sqrt
  assert(excel_sqrt(EXCEL_NUMBER(1)).number == 1.0);
  assert(excel_sqrt(EXCEL_NUMBER(4)).number == 2.0);
  assert(excel_sqrt(EXCEL_NUMBER(1.6)).number == 1.2649110640673518);
  assert(excel_sqrt(BLANK).number == 0);
  assert(excel_sqrt(EXCEL_STRING("Hello world")).type == ExcelError);
  assert(excel_sqrt(EXCEL_NUMBER(-1)).type == ExcelError);

  // Test excel_floor
  assert(excel_floor(EXCEL_NUMBER(1990), EXCEL_NUMBER(100)).number == 1900.0);
  assert(excel_floor(EXCEL_NUMBER(10.99), EXCEL_NUMBER(0.1)).number == 10.9);
  assert(excel_floor(BLANK, ONE).number == 0);
  assert(excel_floor(EXCEL_NUMBER(10.99), ZERO).type == ExcelError);
  assert(excel_floor(EXCEL_NUMBER(10.99), EXCEL_NUMBER(-1.0)).type == ExcelError);
  assert(excel_floor(EXCEL_STRING("Hello world"), ONE).type == ExcelError);
  assert(excel_floor(ONE, EXCEL_STRING("Hello world")).type == ExcelError);
  assert(excel_floor(NA, ONE).type == ExcelError);
  assert(excel_floor(ONE, NA).type == ExcelError);

  // Test partial implementation of rate
  assert(excel_round(multiply(rate(EXCEL_NUMBER(12), ZERO, EXCEL_NUMBER(-69999), EXCEL_NUMBER(64786)), EXCEL_NUMBER(1000)),ONE).number == -6.4);

  // Test MMULT (Matrix multiplication)
  ExcelValue mmult_1[] = { ONE, TWO, THREE, FOUR};
  ExcelValue mmult_2[] = { FOUR, THREE, TWO, ONE};
  ExcelValue mmult_3[] = { ONE, TWO};
  ExcelValue mmult_4[] = { THREE, FOUR};
  ExcelValue mmult_5[] = { ONE, BLANK, THREE, FOUR};

  ExcelValue mmult_1_v = EXCEL_RANGE( mmult_1, 2, 2);
  ExcelValue mmult_2_v = EXCEL_RANGE( mmult_2, 2, 2);
  ExcelValue mmult_3_v = EXCEL_RANGE( mmult_3, 1, 2);
  ExcelValue mmult_4_v = EXCEL_RANGE( mmult_4, 2, 1);
  ExcelValue mmult_5_v = EXCEL_RANGE( mmult_5, 2, 2);

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
  ExcelValue rank_1_a[] = { FIVE, BLANK, THREE, ONE, ONE, FOUR, FIVE, TRUE, SIX, EXCEL_STRING("Hi")};
  ExcelValue rank_2_a[] = { FIVE, BLANK, THREE, NA, ONE, FOUR, FIVE, TRUE, SIX, EXCEL_STRING("Hi")};
  ExcelValue rank_1_v = EXCEL_RANGE( rank_1_a, 2, 5);
  ExcelValue rank_2_v = EXCEL_RANGE( rank_2_a, 2, 5);

  // Basics
  assert(rank(THREE, rank_1_v, ZERO).number == 5);
  assert(rank_2(THREE, rank_1_v).number == 5);
  assert(rank(THREE, rank_1_v, ONE).number == 3);
  assert(rank(ONE, rank_1_v, ZERO).number == 6);
  assert(rank(EXCEL_STRING("3"), rank_1_v, ONE).number == 3);

  // Errors
  assert(rank(TEN, rank_1_v, ZERO).type == ExcelError);
  assert(rank(THREE, rank_2_v, ZERO).type == ExcelError);


  // Test the ISNUMBER function
  assert(excel_isnumber(ONE).type == ExcelBoolean);
  assert(excel_isnumber(ONE).number == 1);
  assert(excel_isnumber(BLANK).type == ExcelBoolean);
  assert(excel_isnumber(BLANK).number == 0);
  assert(excel_isnumber(EXCEL_STRING("Hello")).type == ExcelBoolean);
  assert(excel_isnumber(EXCEL_STRING("Hello")).number == 0);
  assert(excel_isnumber(TRUE).type == ExcelBoolean);
  assert(excel_isnumber(TRUE).number == 0);

  // Test the EXP function
  assert(excel_exp(BLANK).number == 1);
  assert(excel_exp(ZERO).number == 1);
  assert(excel_exp(ONE).number == 2.718281828459045);
  assert(excel_exp(EXCEL_STRING("1")).number == 2.718281828459045);
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
  assert(excel_isblank(EXCEL_STRING("")).number == false);

  // Test AVERAGEIFS function
  ExcelValue averageifs_array_1[] = {EXCEL_NUMBER(10),EXCEL_NUMBER(100),BLANK};
  ExcelValue averageifs_array_1_v = EXCEL_RANGE(averageifs_array_1,3,1);
  ExcelValue averageifs_array_2[] = {EXCEL_STRING("pear"),EXCEL_STRING("bear"),EXCEL_STRING("apple")};
  ExcelValue averageifs_array_2_v = EXCEL_RANGE(averageifs_array_2,3,1);
  ExcelValue averageifs_array_3[] = {EXCEL_NUMBER(1),EXCEL_NUMBER(2),EXCEL_NUMBER(3),EXCEL_NUMBER(4),EXCEL_NUMBER(5),EXCEL_NUMBER(5)};
  ExcelValue averageifs_array_3_v = EXCEL_RANGE(averageifs_array_3,6,1);
  ExcelValue averageifs_array_4[] = {EXCEL_STRING("CO2"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CH4"),EXCEL_STRING("N2O"),EXCEL_STRING("CO2")};
  ExcelValue averageifs_array_4_v = EXCEL_RANGE(averageifs_array_4,6,1);
  ExcelValue averageifs_array_5[] = {EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_STRING("1A"),EXCEL_NUMBER(4),EXCEL_NUMBER(4),EXCEL_NUMBER(5)};
  ExcelValue averageifs_array_5_v = EXCEL_RANGE(averageifs_array_5,6,1);

  // ... should only average values that meet all of the criteria
  ExcelValue averageifs_array_6[] = { averageifs_array_1_v, EXCEL_NUMBER(10), averageifs_array_2_v, EXCEL_STRING("Bear") };
  assert(averageifs(averageifs_array_1_v,4,averageifs_array_6).type == ExcelError);

  ExcelValue averageifs_array_7[] = { averageifs_array_1_v, EXCEL_NUMBER(10), averageifs_array_2_v, EXCEL_STRING("Pear") };
  assert(averageifs(averageifs_array_1_v,4,averageifs_array_7).number == 10.0);

  // ... should work when single cells are given where ranges expected
  ExcelValue averageifs_array_8[] = { EXCEL_STRING("CAR"), EXCEL_STRING("CAR"), EXCEL_STRING("FCV"), EXCEL_STRING("FCV")};
  assert(averageifs(EXCEL_NUMBER(0.143897265452564), 4, averageifs_array_8).number == 0.143897265452564);

  // ... should match numbers with strings that contain numbers
  ExcelValue averageifs_array_9[] = { EXCEL_NUMBER(10), EXCEL_STRING("10.0")};
  assert(averageifs(EXCEL_NUMBER(100),2,averageifs_array_9).number == 100);

  ExcelValue averageifs_array_10[] = { averageifs_array_4_v, EXCEL_STRING("CO2"), averageifs_array_5_v, EXCEL_NUMBER(2)};
  assert(averageifs(averageifs_array_3_v,4, averageifs_array_10).type == ExcelError);

  // ... should match with strings that contain criteria
  ExcelValue averageifs_array_10a[] = { averageifs_array_3_v, EXCEL_STRING("=5")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10a).number == 5);

  ExcelValue averageifs_array_10b[] = { averageifs_array_3_v, EXCEL_STRING("<>3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10b).number == 3.4);

  ExcelValue averageifs_array_10c[] = { averageifs_array_3_v, EXCEL_STRING("<3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10c).number == 1.5);

  ExcelValue averageifs_array_10d[] = { averageifs_array_3_v, EXCEL_STRING("<=3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10d).number == 2);

  ExcelValue averageifs_array_10e[] = { averageifs_array_3_v, EXCEL_STRING(">3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10e).number == 14.0/3.0);

  ExcelValue averageifs_array_10f[] = { averageifs_array_3_v, EXCEL_STRING(">=3")};
  assert(averageifs(averageifs_array_3_v,2, averageifs_array_10f).number == (3.0+4.0+5.0+5.0)/4.0);

  // ... should treat BLANK as an empty string when in the check_range, but not in the criteria
  ExcelValue averageifs_array_11[] = { BLANK, EXCEL_NUMBER(20)};
  assert(averageifs(EXCEL_NUMBER(100),2,averageifs_array_11).type == ExcelError);

  ExcelValue averageifs_array_12[] = {EXCEL_NUMBER(0), BLANK};
  assert(averageifs(EXCEL_NUMBER(100),2,averageifs_array_12).number == 100);

  ExcelValue averageifs_array_13[] = {BLANK, BLANK};
  assert(averageifs(EXCEL_NUMBER(100),2,averageifs_array_13).type == ExcelError);

  // ... should return an error if range argument is an error
  assert(averageifs(REF,2,averageifs_array_13).type == ExcelError);

  // Tests for the FORECAST function
  ExcelValue forecast_array1[] = { EXCEL_NUMBER(1), EXCEL_NUMBER(2), EXCEL_NUMBER(3), EXCEL_NUMBER(4), EXCEL_NUMBER(5)};
  ExcelValue forecast_array2[] = { EXCEL_NUMBER(2), EXCEL_NUMBER(3), EXCEL_NUMBER(4), EXCEL_NUMBER(5), EXCEL_NUMBER(6)};
  ExcelValue forecast_array1_v = EXCEL_RANGE(forecast_array1,5,1);
  ExcelValue forecast_array2_v = EXCEL_RANGE(forecast_array2,5,1);

  assert(forecast(EXCEL_NUMBER(0), forecast_array2_v, forecast_array1_v).number == 1);
  assert(forecast(EXCEL_NUMBER(1), forecast_array2_v, forecast_array1_v).number == 2);
  assert(forecast(EXCEL_NUMBER(6), forecast_array2_v, forecast_array1_v).number == 7);

  ExcelValue forecast_array3[] = { BLANK, EXCEL_NUMBER(2), EXCEL_NUMBER(3), EXCEL_NUMBER(4), BLANK};
  ExcelValue forecast_array3_v = EXCEL_RANGE(forecast_array3,5,1);

  assert(forecast(EXCEL_NUMBER(6), forecast_array2_v, forecast_array3_v).number == 7);

  // Tests ENSURE_IS_NUMBER function
  assert(ensure_is_number(EXCEL_NUMBER(1.3)).type == ExcelNumber);
  assert(ensure_is_number(EXCEL_NUMBER(1.3)).number == 1.3);
  assert(ensure_is_number(BLANK).type == ExcelNumber);
  assert(ensure_is_number(BLANK).number == 0);
  assert(ensure_is_number(TRUE).type == ExcelNumber);
  assert(ensure_is_number(TRUE).number == 1.0);
  assert(ensure_is_number(FALSE).type == ExcelNumber);
  assert(ensure_is_number(FALSE).number == 0.0);
  assert(ensure_is_number(EXCEL_STRING("1.3")).type == ExcelNumber);
  assert(ensure_is_number(EXCEL_STRING("1.3")).number == 1.3);
  assert(ensure_is_number(EXCEL_STRING("BASDASD")).type == ExcelError);
  assert(ensure_is_number(DIV0).type == ExcelError);

  // Tests ther NUMBER_OR_ZERO function
  assert_equal(ZERO, number_or_zero(ZERO), "number_or_zero 0");
  assert_equal(ONE, number_or_zero(ONE), "number_or_zero 1");
  assert_equal(VALUE, number_or_zero(VALUE), "number_or_zero :error");
  assert_equal(ZERO, number_or_zero(TRUE), "number_or_zero true");
  assert_equal(ZERO, number_or_zero(FALSE), "number_or_zero false");
  assert_equal(ZERO, number_or_zero(BLANK), "number_or_zero blank");
  assert_equal(ZERO, number_or_zero(EXCEL_STRING("1.3")), "number_or_zero '1.3'");
  assert_equal(ZERO, number_or_zero(EXCEL_STRING("Aasdfadsf")), "number_or_zero 'Asdfad'");

  // RIGHT(string,[characters])
  // ... should return the right n characters from a string
  assert(strcmp(right_1(EXCEL_STRING("ONE")).string,"E") == 0);
  assert(strcmp(right(EXCEL_STRING("ONE"),ONE).string,"E") == 0);
  assert(strcmp(right(EXCEL_STRING("ONE"),EXCEL_NUMBER(3)).string,"ONE") == 0);
  // ... should turn numbers into strings before processing
  assert(strcmp(right(EXCEL_NUMBER(1.31e12),EXCEL_NUMBER(3)).string, "000") == 0);
  // ... should turn booleans into the words TRUE and FALSE before processing
  assert(strcmp(right(TRUE,EXCEL_NUMBER(3)).string,"RUE") == 0);
  assert(strcmp(right(FALSE,EXCEL_NUMBER(3)).string,"LSE") == 0);
  // ... should return BLANK if given BLANK for either argument
  assert(right(BLANK,EXCEL_NUMBER(3)).type == ExcelEmpty);
  assert(right(EXCEL_STRING("ONE"),BLANK).type == ExcelEmpty);
  // ... should return an error if an argument is an error
  assert(right_1(NA).type == ExcelError);
  assert(right(EXCEL_STRING("ONE"),NA).type == ExcelError);
  assert(right(EXCEL_STRING("ONE"),EXCEL_NUMBER(-10)).type == ExcelError);
  assert_equal(EXCEL_STRING("ONE"), right(EXCEL_STRING("ONE"), EXCEL_NUMBER(100)), "RIGHT if number of characters greater than string length");

  // LEN(string)
  assert(len(BLANK).type == ExcelNumber);
  assert(len(BLANK).number == 0);
  assert(len(EXCEL_STRING("Hello")).number == 5);
  assert(len(EXCEL_NUMBER(123)).number == 3);
  assert(len(TRUE).number == 4);
  assert(len(FALSE).number == 5);


  // VALUE(something)
  assert(value(BLANK).type == ExcelNumber);
  assert(value(BLANK).number == 0);
  assert(value(ONE).type == ExcelNumber);
  assert(value(ONE).number == 1);
  assert(value(EXCEL_STRING("1")).type == ExcelNumber);
  assert(value(EXCEL_STRING("1")).number == 1);
  assert(value(EXCEL_STRING("A1A")).type == ExcelError);


  // NPV(rate, flow1, flow2)
  ExcelValue npv_array1[] = { EXCEL_NUMBER(110) };
  assert(npv(EXCEL_NUMBER(0.1), 1, npv_array1).type == ExcelNumber);
  assert(npv(EXCEL_NUMBER(0.1), 1, npv_array1).number-100 < 0.001);

  ExcelValue npv_array2[] = { EXCEL_NUMBER(110), EXCEL_NUMBER(121) };
  assert((npv(EXCEL_NUMBER(0.1), 2, npv_array2).number - 200) < 0.001);

  ExcelValue npv_array3[] = { EXCEL_NUMBER(110), EXCEL_NUMBER(121)};
  ExcelValue npv_array3_v = EXCEL_RANGE(npv_array3,2,1);
  ExcelValue npv_array4[] = { npv_array3_v };

  assert((npv(EXCEL_NUMBER(0.1), 1,  npv_array4).number - 200) < 0.001);

  assert(npv(EXCEL_NUMBER(-1.0), 1, npv_array1).type == ExcelError);
  assert(npv(BLANK, 1, npv_array1).number == 110);

  ExcelValue npv_array5[] = { BLANK };
  assert(npv(EXCEL_NUMBER(0.1), 1, npv_array5).number == 0);

  // CHAR
  assert_equal(VALUE, excel_char(ZERO), "excel_char(0) == VALUE");
  assert_equal(VALUE, excel_char(EXCEL_NUMBER(256)), "excel_char(256) == VALUE");
  assert_equal(VALUE, excel_char(BLANK), "excel_char() == VALUE");
  assert_equal(VALUE, excel_char(EXCEL_STRING("adsfa")), "excel_char('nonsense') == VALUE");
  assert_equal(DIV0, excel_char(DIV0), "excel_char(DIV0) == DIV0");
  assert_equal(EXCEL_STRING("\x01"), excel_char(ONE), "excel_char(1) == '\x01'");
  assert_equal(EXCEL_STRING("a"), excel_char(EXCEL_NUMBER(97)), "excel_char(97) == 'a'");
  assert_equal(EXCEL_STRING("a"), excel_char(EXCEL_NUMBER(97.5)), "excel_char(97.5) == 'a'");

  // NA()
  assert_equal(NA, na(), "na() == NA");

  // curve (a custom climact function)
  assert_equal(
    EXCEL_NUMBER(4.99),
    curve_5(
      EXCEL_STRING("s"),
      EXCEL_NUMBER(2023),
      EXCEL_NUMBER(0),
      EXCEL_NUMBER(10),
      EXCEL_NUMBER(10)
    ),
    "curve_5('s', 2023, 0, 10, 10) == 4.99"
  );

  // If blank, defaults to lcurve
  assert_equal(
    EXCEL_NUMBER(5.0),
    curve_5(
      BLANK,
      EXCEL_NUMBER(2023),
      EXCEL_NUMBER(0),
      EXCEL_NUMBER(10),
      EXCEL_NUMBER(10)
    ),
    "curve_5(blank, 2023, 0, 10, 10) == 5"
  );


  // Release memory
  free_all_allocated_memory();

  // Yay!
  printf("\nFinished tests\n");

  return 0;
}

int main() {
  return test_functions();
}
