# coding: utf-8
# Test for singlefunctionc
require 'minitest/autorun'
require_relative 'singlefunctionc'

class TestSinglefunctionc < Minitest::Unit::TestCase
  def test_run
    workbook = Singlefunctionc.new

    workbook.input = 100.0
    workbook.inputs = [[10.0, 20.0, 30.0, 40.0]]

    assert_equal(workbook.output, 0.9900498337491681)
    assert_equal(workbook.outputs, [[-9.024690087971667, 0.951229424500714, 0.9672161004820059, 0.9753099120283326]])
    assert_equal(workbook.input, 100.0)
    assert_equal(workbook.inputs, [[10.0, 20.0, 30.0, 40.0]])

  end
end
