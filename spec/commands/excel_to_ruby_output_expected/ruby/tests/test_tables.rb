# Test for Tables
require 'test/unit'
require_relative 'tables'

module ExampleSpreadsheet
class TestTables < Test::Unit::TestCase
  def test_b2; assert_equal(worksheet.b2,"ColA"); end
  def test_c2; assert_equal(worksheet.c2,"ColB"); end
  def test_d2; assert_equal(worksheet.d2,"Column1"); end
  def test_g2; assert_equal(worksheet.g2,:value); end
  def test_b3; assert_equal(worksheet.b3,1); end
  def test_c3; assert_equal(worksheet.c3,"A"); end
  def test_d3; assert_equal(worksheet.d3,:value); end
  def test_b4; assert_equal(worksheet.b4,2); end
  def test_c4; assert_equal(worksheet.c4,"B"); end
  def test_d4; assert_equal(worksheet.d4,:value); end
  def test_f4; assert_equal(worksheet.f4,"B"); end
  def test_g4; assert_equal(worksheet.g4,:value); end
  def test_h4; assert_equal(worksheet.h4,:value); end
  def test_b5; assert_equal(worksheet.b5,3); end
  def test_c5; assert_equal(worksheet.c5,0); end
  def test_e6; assert_equal(worksheet.e6,:value); end
  def test_e7; assert_equal(worksheet.e7,:value); end
  def test_e8; assert_equal(worksheet.e8,:value); end
  def test_e9; assert_equal(worksheet.e9,:value); end
  def test_c10; assert_equal(worksheet.c10,0); end
  def test_e10; assert_equal(worksheet.e10,:value); end
  def test_c11; assert_equal(worksheet.c11,3); end
  def test_c12; assert_equal(worksheet.c12,3); end
end
end
