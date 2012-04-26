# coding: utf-8
# Test for blank
require 'rubygems'
gem 'minitest'
require 'test/unit'
require_relative 'blank'

class TestBlank < Test::Unit::TestCase
  def spreadsheet; @spreadsheet ||= init_spreadsheet; end
  def init_spreadsheet; Blank end

  # start of Sheet1
end
