# coding: utf-8
# Test for utf8strings
require 'rubygems'
gem 'minitest'
require 'minitest/autorun'
require_relative 'utf8strings'

class TestUtf8strings < Minitest::Test
  def self.runnable_methods
    puts 'Overriding minitest to run tests in a defined order'
    methods = methods_matching(/^test_/)
  end
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; Utf8stringsShim.new end
  def test_s2015_b3; assert_equal("連動", worksheet.s2015_b3); end
  def test_sheet1_a1; assert_equal("string", worksheet.sheet1_a1); end
  def test_sheet1_b1; assert_equal("hello world", worksheet.sheet1_b1); end
  def test_sheet1_a2; assert_equal("Number", worksheet.sheet1_a2); end
  def test_sheet1_b2; assert_in_epsilon(3.0, worksheet.sheet1_b2, 0.002); end
  def test_sheet1_a3; assert_equal("Result", worksheet.sheet1_a3); end
  def test_sheet1_b3; assert_equal("hello world3", worksheet.sheet1_b3); end
  def test_s_1_a2; assert_equal("test…3", worksheet.s_1_a2); end
  def test_s_1_c2; assert_equal("test…3test…", worksheet.s_1_c2); end
  def test_s_1_b4; assert_equal("test…", worksheet.s_1_b4); end
  def test_s_1_b5; assert_in_epsilon(3.0, worksheet.s_1_b5, 0.002); end
  def test_s_1_b6; assert_equal("測試", worksheet.s_1_b6); end
  def test_s_1_b9; assert_in_epsilon(2015.0, worksheet.s_1_b9, 0.002); end
  def test_s_1_b11; assert_equal("連動", worksheet.s_1_b11); end
end
