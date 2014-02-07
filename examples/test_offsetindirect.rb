# coding: utf-8
# Test for offsetindirect
require 'minitest/autorun'
require_relative 'offsetindirect'

class TestOffsetindirect < Minitest::Unit::TestCase
  def self.runnable_methods
    puts 'Overriding minitest to run tests in a defined order'
    methods = methods_matching(/^test_/)
  end
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; OffsetindirectShim.new end
  def test_sheet1_a7; assert_equal("G.30", worksheet.sheet1_a7); end
  def test_sheet1_b7; assert_equal("Name for G.30", worksheet.sheet1_b7); end
  def test_sheet1_c7; assert_equal("G.30.Choice", worksheet.sheet1_c7); end
  def test_sheet1_a9; assert_in_epsilon(30.0, worksheet.sheet1_a9, 0.002); end
  def test_sheet1_b9; assert_equal("Name for G.30", worksheet.sheet1_b9); end
end
