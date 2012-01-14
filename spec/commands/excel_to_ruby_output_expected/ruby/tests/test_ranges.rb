# Test for Ranges
require 'test/unit'
require_relative 'ranges'

module ExampleSpreadsheet
class TestRanges < Test::Unit::TestCase
  def test_b1; assert_equal(worksheet.b1,"This sheet"); end
  def test_c1; assert_equal(worksheet.c1,"Other sheet"); end
  def test_a2; assert_equal(worksheet.a2,"Standard"); end
  def test_b2; assert_equal(worksheet.b2,6); end
  def test_c2; assert_equal(worksheet.c2,4.141500000000001); end
  def test_a3; assert_equal(worksheet.a3,"Column"); end
  def test_b3; assert_equal(worksheet.b3,6); end
  def test_c3; assert_equal(worksheet.c3,:name); end
  def test_a4; assert_equal(worksheet.a4,"Row"); end
  def test_b4; assert_equal(worksheet.b4,6); end
  def test_c4; assert_equal(worksheet.c4,3.1415); end
  def test_f4; assert_equal(worksheet.f4,1); end
  def test_e5; assert_equal(worksheet.e5,1); end
  def test_f5; assert_equal(worksheet.f5,2); end
  def test_g5; assert_equal(worksheet.g5,3); end
  def test_f6; assert_equal(worksheet.f6,3); end
end
end
