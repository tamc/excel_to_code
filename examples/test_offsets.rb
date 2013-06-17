# coding: utf-8
# Test for offsets
require 'rubygems'
gem 'minitest'
require 'test/unit'
require_relative 'offsets'

class TestOffsets < Test::Unit::TestCase
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; OffsetsShim.new end

  # start of Model
  def test_model_c22; assert_in_delta(-0.004, worksheet.model_c22, 0.001); end
  def test_model_c23; assert_in_epsilon(2020, worksheet.model_c23, 0.001); end
  def test_model_c24; assert_in_epsilon(600, worksheet.model_c24, 0.001); end
  def test_model_c25; assert_in_delta(0.5, worksheet.model_c25, 0.001); end
  def test_model_c26; assert_in_epsilon(2, worksheet.model_c26, 0.001); end
  def test_model_c29; assert_in_epsilon(2.5, worksheet.model_c29, 0.001); end
  def test_model_c30; assert_in_delta(0.65, worksheet.model_c30, 0.001); end
  def test_model_c31; assert_in_epsilon(11, worksheet.model_c31, 0.001); end
  def test_model_c33; assert_in_epsilon(2030, worksheet.model_c33, 0.001); end
  def test_model_c34; assert_in_epsilon(5, worksheet.model_c34, 0.001); end
  def test_model_c35; assert_in_delta(0.4, worksheet.model_c35, 0.001); end
  def test_model_c37; assert_in_delta(1, worksheet.model_c37, 0.001); end
  def test_model_c38; assert_in_delta(1, worksheet.model_c38, 0.001); end
  def test_model_c39; assert_in_delta(0.6, worksheet.model_c39, 0.001); end
  def test_model_c41; assert_in_epsilon(60, worksheet.model_c41, 0.001); end
  def test_model_c42; assert_in_epsilon(60, worksheet.model_c42, 0.001); end
  def test_model_c43; assert_in_epsilon(35, worksheet.model_c43, 0.001); end
  def test_model_c46; assert_in_epsilon(432, worksheet.model_c46, 0.001); end
  def test_model_c47; assert_in_epsilon(350, worksheet.model_c47, 0.001); end
  def test_model_c48; assert_in_delta(0, (worksheet.model_c48||0), 0.001); end
  def test_model_c76; assert_in_epsilon(8.766, worksheet.model_c76, 0.001); end
  def test_model_c77; assert_in_delta(0.11407711613050422, worksheet.model_c77, 0.001); end
  def test_model_c107; assert_in_epsilon(1.625, worksheet.model_c107, 0.001); end
  def test_model_c108; assert_in_delta(0.875, worksheet.model_c108, 0.001); end
  def test_model_c109; assert_in_epsilon(2.5, worksheet.model_c109, 0.001); end
  def test_model_c110; assert_in_delta(0, (worksheet.model_c110||0), 0.001); end
  def test_model_c113; assert_in_epsilon(7, worksheet.model_c113, 0.001); end
  def test_model_c114; assert_in_epsilon(15.292, worksheet.model_c114, 0.001); end
  def test_model_c115; assert_in_epsilon(22.292, worksheet.model_c115, 0.001); end
  def test_model_c116; assert_in_epsilon(66.82, worksheet.model_c116, 0.001); end
  def test_model_c117; assert_in_epsilon(89.112, worksheet.model_c117, 0.001); end
  def test_model_c119; assert_in_epsilon(20.627044196734136, worksheet.model_c119, 0.001); end
  def test_model_c120; assert_in_epsilon(40.352498288843265, worksheet.model_c120, 0.001); end
  def test_model_c121; assert_in_epsilon(63.72685714285714, worksheet.model_c121, 0.001); end
  def test_model_c124; assert_in_delta(0.25667351129363447, worksheet.model_c124, 0.001); end
  def test_model_c126; assert_in_delta(0.7534104076659957, worksheet.model_c126, 0.001); end
  def test_model_c128; assert_in_delta(0.6332102365943555, worksheet.model_c128, 0.001); end
  def test_model_c131; assert_in_epsilon(15.749999999999998, worksheet.model_c131, 0.001); end
  def test_model_c132; assert_in_epsilon(131.47518746421383, worksheet.model_c132, 0.001); end
  def test_model_c133; assert_in_epsilon(147.22518746421383, worksheet.model_c133, 0.001); end
  def test_model_c137; assert_in_epsilon(443.83783567231035, worksheet.model_c137, 0.001); end
  def test_model_c139; assert_in_epsilon(156.99875761236635, worksheet.model_c139, 0.001); end
  def test_model_c140; assert_in_epsilon(78, worksheet.model_c140, 0.001); end
  def test_model_c141; assert_in_epsilon(234.99875761236635, worksheet.model_c141, 0.001); end
  def test_model_c142; assert_in_epsilon(246.6, worksheet.model_c142, 0.001); end
  def test_model_c143; assert_in_epsilon(246.6, worksheet.model_c143, 0.001); end
  def test_model_c1243; assert_in_delta(0.25667351129363447, worksheet.model_c1243, 0.001); end
  def test_model_c1244; assert_in_delta(0.9807945480404743, worksheet.model_c1244, 0.001); end
  def test_model_c1245; assert_in_delta(0.3525512343782234, worksheet.model_c1245, 0.001); end
  def test_model_c1248; assert_in_epsilon(15.749999999999998, worksheet.model_c1248, 0.001); end
  def test_model_c1249; assert_in_epsilon(131.47518746421383, worksheet.model_c1249, 0.001); end
  def test_model_c1250; assert_in_epsilon(206.50481253578616, worksheet.model_c1250, 0.001); end
  def test_model_c1251; assert_in_epsilon(353.73, worksheet.model_c1251, 0.001); end
  def test_model_c1252; assert_in_delta(0, (worksheet.model_c1252||0), 0.001); end
end
