# coding: utf-8
# Test for offsets
require 'rubygems'
gem 'minitest'
require 'test/unit'
require_relative 'offsets'

class TestOffsets < Test::Unit::TestCase
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; OffsetsShim.new end
end
