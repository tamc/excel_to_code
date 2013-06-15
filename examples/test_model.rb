# coding: utf-8
# Test for model
require 'rubygems'
gem 'minitest'
require 'test/unit'
require_relative 'model'

class TestModel < Test::Unit::TestCase
  def worksheet; @worksheet ||= init_spreadsheet; end
  def init_spreadsheet; ModelShim.new end
end
