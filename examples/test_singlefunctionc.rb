# coding: utf-8
# Test for singlefunctionc
require 'minitest/autorun'
require_relative 'singlefunctionc'

class TestSinglefunctionc < Minitest::Unit::TestCase
  def test_run
    workbook = Singlefunctionc.new

    workbook.input = 100.0

    assert_equal(workbook.output, 0.9900498337491681)
    assert_equal(workbook.input, 100.0)

  end
end
