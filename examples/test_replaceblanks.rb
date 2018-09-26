# coding: utf-8
# Test for replaceblanks
require 'minitest/autorun'
require_relative 'replaceblanks'

class TestReplaceblanks < Minitest::Unit::TestCase
  def self.runnable_methods
    puts 'Overriding minitest to run tests in a defined order'
    methods = methods_matching(/^test_/)
  end
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; Replaceblanks.new end
  def test_sheet1_c4; assert_in_epsilon(38.0, worksheet.sheet1_c4, 0.002); end
  def test_sheet1_c5; assert_in_delta(0.0, (worksheet.sheet1_c5||0), 0.002); end
  def test_sheet2_g19; assert_in_epsilon(38.0, worksheet.sheet2_g19, 0.002); end
  def test_sheet2_g20; assert_in_delta(0.0, (worksheet.sheet2_g20||0), 0.002); end
  def test_sheet2_g21; assert_in_delta(0.0, (worksheet.sheet2_g21||0), 0.002); end
end
