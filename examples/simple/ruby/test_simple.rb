# coding: utf-8
# All tests for /Users/tamc/Documents/github/excel_to_code/examples/simple/simple.xlsx
require 'minitest/autorun'
require_relative 'simple'

class TestSimple < Minitest::Unit::TestCase
  def worksheet; @worksheet ||= Simple.new; end
  def test_inputs_a1; assert_equal("Your first name:", worksheet.inputs_a1.to_s.gsub(/[\n\r]+/,'')); end
  def test_inputs_b1; assert_equal("Harry", worksheet.inputs_b1.to_s.gsub(/[\n\r]+/,'')); end
  def test_inputs_a2; assert_equal("Your surname:", worksheet.inputs_a2.to_s.gsub(/[\n\r]+/,'')); end
  def test_inputs_b2; assert_equal("Potter", worksheet.inputs_b2.to_s.gsub(/[\n\r]+/,'')); end
  def test_inputs_a4; assert_equal("What is the meaning of life?", worksheet.inputs_a4.to_s.gsub(/[\n\r]+/,'')); end
  def test_inputs_b4; assert_in_epsilon(24.0, worksheet.inputs_b4, 0.002); end
  def test_outputs_a1; assert_equal("Hello Harry Potter", worksheet.outputs_a1.to_s.gsub(/[\n\r]+/,'')); end
  def test_outputs_a2; assert_equal("Oh dear, that isn't the meaning of life. Guess again.", worksheet.outputs_a2.to_s.gsub(/[\n\r]+/,'')); end
  def test_outputs_a3; assert_equal("You were 18 out.", worksheet.outputs_a3.to_s.gsub(/[\n\r]+/,'')); end
end
