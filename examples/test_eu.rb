# coding: utf-8
# Test for eu
require 'rubygems'
gem 'minitest'
require 'test/unit'
require_relative 'eu'

class TestEu < Test::Unit::TestCase
  def spreadsheet; @spreadsheet ||= init_spreadsheet; end
  def init_spreadsheet; Eu end

  # start of EU
def test_eu_b1
  r = spreadsheet.eu_b1
  assert_equal(:ExcelString,r[:type])
  assert_equal("Note, numbers not checked. Worry about power sector figure. Worry about the UK share of auction revenues.",r[:string].force_encoding('utf-8'))
end

def test_eu_d2
  r = spreadsheet.eu_d2
  assert_equal(:ExcelString,r[:type])
  assert_equal("Expansion",r[:string].force_encoding('utf-8'))
end

def test_eu_d3
  r = spreadsheet.eu_d3
  assert_equal(:ExcelString,r[:type])
  assert_equal("Additions",r[:string].force_encoding('utf-8'))
end

def test_eu_o3
  r = spreadsheet.eu_o3
  assert_equal(:ExcelString,r[:type])
  assert_equal("Annual",r[:string].force_encoding('utf-8'))
end

def test_eu_b4
  r = spreadsheet.eu_b4
  assert_equal(:ExcelString,r[:type])
  assert_equal("EU-27 Emissions",r[:string].force_encoding('utf-8'))
end

def test_eu_c4
  r = spreadsheet.eu_c4
  assert_equal(:ExcelString,r[:type])
  assert_equal("2005-6",r[:string].force_encoding('utf-8'))
end

def test_eu_d4
  r = spreadsheet.eu_d4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Phase II",r[:string].force_encoding('utf-8'))
end

def test_eu_e4
  r = spreadsheet.eu_e4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Phase III",r[:string].force_encoding('utf-8'))
end

def test_eu_f4
  r = spreadsheet.eu_f4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g4
  r = spreadsheet.eu_g4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h4
  r = spreadsheet.eu_h4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i4
  r = spreadsheet.eu_i4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j4
  r = spreadsheet.eu_j4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k4
  r = spreadsheet.eu_k4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l4
  r = spreadsheet.eu_l4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m4
  r = spreadsheet.eu_m4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_o4
  r = spreadsheet.eu_o4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Change",r[:string].force_encoding('utf-8'))
end

def test_eu_b5
  r = spreadsheet.eu_b5
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_eu_c5
  r = spreadsheet.eu_c5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1150.0,r[:number])
end

def test_eu_d5
  r = spreadsheet.eu_d5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(50.0,r[:number])
end

def test_eu_e5
  r = spreadsheet.eu_e5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(50.0,r[:number])
end

def test_eu_f5
  r = spreadsheet.eu_f5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1125.0,r[:number])
end

def test_eu_g5
  r = spreadsheet.eu_g5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1102.5,r[:number])
end

def test_eu_h5
  r = spreadsheet.eu_h5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1080.45,r[:number])
end

def test_eu_i5
  r = spreadsheet.eu_i5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1058.8410000000001,r[:number])
end

def test_eu_j5
  r = spreadsheet.eu_j5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1037.66418,r[:number])
end

def test_eu_k5
  r = spreadsheet.eu_k5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1016.9108964,r[:number])
end

def test_eu_l5
  r = spreadsheet.eu_l5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(996.5726784719999,r[:number])
end

def test_eu_m5
  r = spreadsheet.eu_m5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(976.6412249025599,r[:number])
end

def test_eu_n5
  r = spreadsheet.eu_n5
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_o5
  r = spreadsheet.eu_o5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.02,r[:number])
end

def test_eu_b6
  r = spreadsheet.eu_b6
  assert_equal(:ExcelString,r[:type])
  assert_equal("Leakage sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_c6
  r = spreadsheet.eu_c6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(350.0,r[:number])
end

def test_eu_f6
  r = spreadsheet.eu_f6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(332.5,r[:number])
end

def test_eu_g6
  r = spreadsheet.eu_g6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(325.84999999999997,r[:number])
end

def test_eu_h6
  r = spreadsheet.eu_h6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(319.33299999999997,r[:number])
end

def test_eu_i6
  r = spreadsheet.eu_i6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(312.94633999999996,r[:number])
end

def test_eu_j6
  r = spreadsheet.eu_j6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(306.6874132,r[:number])
end

def test_eu_k6
  r = spreadsheet.eu_k6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(300.55366493599996,r[:number])
end

def test_eu_l6
  r = spreadsheet.eu_l6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(294.54259163727994,r[:number])
end

def test_eu_m6
  r = spreadsheet.eu_m6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(288.6517398045343,r[:number])
end

def test_eu_n6
  r = spreadsheet.eu_n6
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_o6
  r = spreadsheet.eu_o6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.01,r[:number])
end

def test_eu_b7
  r = spreadsheet.eu_b7
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_c7
  r = spreadsheet.eu_c7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(550.0,r[:number])
end

def test_eu_d7
  r = spreadsheet.eu_d7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(100.0,r[:number])
end

def test_eu_e7
  r = spreadsheet.eu_e7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(100.0,r[:number])
end

def test_eu_f7
  r = spreadsheet.eu_f7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(712.5,r[:number])
end

def test_eu_g7
  r = spreadsheet.eu_g7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(698.25,r[:number])
end

def test_eu_h7
  r = spreadsheet.eu_h7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(684.285,r[:number])
end

def test_eu_i7
  r = spreadsheet.eu_i7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(670.5993,r[:number])
end

def test_eu_j7
  r = spreadsheet.eu_j7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(657.187314,r[:number])
end

def test_eu_k7
  r = spreadsheet.eu_k7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(644.04356772,r[:number])
end

def test_eu_l7
  r = spreadsheet.eu_l7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(631.1626963656,r[:number])
end

def test_eu_m7
  r = spreadsheet.eu_m7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(618.539442438288,r[:number])
end

def test_eu_n7
  r = spreadsheet.eu_n7
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_o7
  r = spreadsheet.eu_o7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.01,r[:number])
end

def test_eu_b8
  r = spreadsheet.eu_b8
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_eu_f8
  r = spreadsheet.eu_f8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2170.0,r[:number])
end

def test_eu_g8
  r = spreadsheet.eu_g8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2126.6,r[:number])
end

def test_eu_h8
  r = spreadsheet.eu_h8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2084.0679999999998,r[:number])
end

def test_eu_i8
  r = spreadsheet.eu_i8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2042.3866400000002,r[:number])
end

def test_eu_j8
  r = spreadsheet.eu_j8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2001.5389072,r[:number])
end

def test_eu_k8
  r = spreadsheet.eu_k8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1961.508129056,r[:number])
end

def test_eu_l8
  r = spreadsheet.eu_l8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1922.27796647488,r[:number])
end

def test_eu_m8
  r = spreadsheet.eu_m8
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1883.8324071453821,r[:number])
end

def test_eu_n8
  r = spreadsheet.eu_n8
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_b10
  r = spreadsheet.eu_b10
  assert_equal(:ExcelString,r[:type])
  assert_equal("Proportion allocated for free",r[:string].force_encoding('utf-8'))
end

def test_eu_f10
  r = spreadsheet.eu_f10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g10
  r = spreadsheet.eu_g10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h10
  r = spreadsheet.eu_h10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i10
  r = spreadsheet.eu_i10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j10
  r = spreadsheet.eu_j10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k10
  r = spreadsheet.eu_k10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l10
  r = spreadsheet.eu_l10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m10
  r = spreadsheet.eu_m10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_b11
  r = spreadsheet.eu_b11
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_eu_f11
  r = spreadsheet.eu_f11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_g11
  r = spreadsheet.eu_g11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_h11
  r = spreadsheet.eu_h11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_i11
  r = spreadsheet.eu_i11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_j11
  r = spreadsheet.eu_j11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_k11
  r = spreadsheet.eu_k11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_l11
  r = spreadsheet.eu_l11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_m11
  r = spreadsheet.eu_m11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_n11
  r = spreadsheet.eu_n11
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_eu_b12
  r = spreadsheet.eu_b12
  assert_equal(:ExcelString,r[:type])
  assert_equal("Leakage sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_f12
  r = spreadsheet.eu_f12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.9,r[:number])
end

def test_eu_g12
  r = spreadsheet.eu_g12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.875,r[:number])
end

def test_eu_h12
  r = spreadsheet.eu_h12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.85,r[:number])
end

def test_eu_i12
  r = spreadsheet.eu_i12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.825,r[:number])
end

def test_eu_j12
  r = spreadsheet.eu_j12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7999999999999999,r[:number])
end

def test_eu_k12
  r = spreadsheet.eu_k12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7749999999999999,r[:number])
end

def test_eu_l12
  r = spreadsheet.eu_l12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7499999999999999,r[:number])
end

def test_eu_m12
  r = spreadsheet.eu_m12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7249999999999999,r[:number])
end

def test_eu_o12
  r = spreadsheet.eu_o12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(-0.025,r[:number])
end

def test_eu_b13
  r = spreadsheet.eu_b13
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other",r[:string].force_encoding('utf-8'))
end

def test_eu_f13
  r = spreadsheet.eu_f13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8,r[:number])
end

def test_eu_g13
  r = spreadsheet.eu_g13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.6857142857142857,r[:number])
end

def test_eu_h13
  r = spreadsheet.eu_h13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5714285714285714,r[:number])
end

def test_eu_i13
  r = spreadsheet.eu_i13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.4571428571428571,r[:number])
end

def test_eu_j13
  r = spreadsheet.eu_j13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.34285714285714275,r[:number])
end

def test_eu_k13
  r = spreadsheet.eu_k13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.22857142857142845,r[:number])
end

def test_eu_l13
  r = spreadsheet.eu_l13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.11428571428571416,r[:number])
end

def test_eu_m13
  r = spreadsheet.eu_m13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0,r[:number])
end

def test_eu_n13
  r = spreadsheet.eu_n13
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_eu_o13
  r = spreadsheet.eu_o13
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(-0.1142857142857143,r[:number])
end

def test_eu_b14
  r = spreadsheet.eu_b14
  assert_equal(:ExcelString,r[:type])
  assert_equal("Average non-power sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_f14
  r = spreadsheet.eu_f14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8318181818181818,r[:number])
end

def test_eu_g14
  r = spreadsheet.eu_g14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7459415584415585,r[:number])
end

def test_eu_h14
  r = spreadsheet.eu_h14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.660064935064935,r[:number])
end

def test_eu_i14
  r = spreadsheet.eu_i14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5741883116883116,r[:number])
end

def test_eu_j14
  r = spreadsheet.eu_j14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.4883116883116882,r[:number])
end

def test_eu_k14
  r = spreadsheet.eu_k14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.40243506493506487,r[:number])
end

def test_eu_l14
  r = spreadsheet.eu_l14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.3165584415584414,r[:number])
end

def test_eu_m14
  r = spreadsheet.eu_m14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.23068181818181807,r[:number])
end

def test_eu_b15
  r = spreadsheet.eu_b15
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total free allocation, % emissions",r[:string].force_encoding('utf-8'))
end

def test_eu_f15
  r = spreadsheet.eu_f15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.40057603686635945,r[:number])
end

def test_eu_g15
  r = spreadsheet.eu_g15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.3592207044107966,r[:number])
end

def test_eu_h15
  r = spreadsheet.eu_h15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.31786537195523373,r[:number])
end

def test_eu_i15
  r = spreadsheet.eu_i15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.2765100394996708,r[:number])
end

def test_eu_j15
  r = spreadsheet.eu_j15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.23515470704410793,r[:number])
end

def test_eu_k15
  r = spreadsheet.eu_k15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.19379937458854504,r[:number])
end

def test_eu_l15
  r = spreadsheet.eu_l15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.15244404213298213,r[:number])
end

def test_eu_m15
  r = spreadsheet.eu_m15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.11108870967741932,r[:number])
end

def test_eu_n15
  r = spreadsheet.eu_n15
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_eu_b17
  r = spreadsheet.eu_b17
  assert_equal(:ExcelString,r[:type])
  assert_equal("Proportion auctioned, other sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_f17
  r = spreadsheet.eu_f17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.1681818181818182,r[:number])
end

def test_eu_g17
  r = spreadsheet.eu_g17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.2540584415584415,r[:number])
end

def test_eu_h17
  r = spreadsheet.eu_h17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.339935064935065,r[:number])
end

def test_eu_i17
  r = spreadsheet.eu_i17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.42581168831168836,r[:number])
end

def test_eu_j17
  r = spreadsheet.eu_j17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5116883116883117,r[:number])
end

def test_eu_k17
  r = spreadsheet.eu_k17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5975649350649351,r[:number])
end

def test_eu_l17
  r = spreadsheet.eu_l17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.6834415584415586,r[:number])
end

def test_eu_m17
  r = spreadsheet.eu_m17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7693181818181819,r[:number])
end

def test_eu_b18
  r = spreadsheet.eu_b18
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total allowances",r[:string].force_encoding('utf-8'))
end

def test_eu_d18
  r = spreadsheet.eu_d18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2010.0,r[:number])
end

def test_eu_f18
  r = spreadsheet.eu_f18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g18
  r = spreadsheet.eu_g18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h18
  r = spreadsheet.eu_h18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i18
  r = spreadsheet.eu_i18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j18
  r = spreadsheet.eu_j18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k18
  r = spreadsheet.eu_k18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l18
  r = spreadsheet.eu_l18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m18
  r = spreadsheet.eu_m18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_b19
  r = spreadsheet.eu_b19
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total allowances",r[:string].force_encoding('utf-8'))
end

def test_eu_d19
  r = spreadsheet.eu_d19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2207.0,r[:number])
end

def test_eu_f19
  r = spreadsheet.eu_f19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2091.7946,r[:number])
end

def test_eu_g19
  r = spreadsheet.eu_g19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2055.39737396,r[:number])
end

def test_eu_h19
  r = spreadsheet.eu_h19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.6334596530962,r[:number])
end

def test_eu_i19
  r = spreadsheet.eu_i19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1984.4918374551323,r[:number])
end

def test_eu_j19
  r = spreadsheet.eu_j19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1949.961679483413,r[:number])
end

def test_eu_k19
  r = spreadsheet.eu_k19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1916.0323462604017,r[:number])
end

def test_eu_l19
  r = spreadsheet.eu_l19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1882.6933834354709,r[:number])
end

def test_eu_m19
  r = spreadsheet.eu_m19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1849.9345185636937,r[:number])
end

def test_eu_o19
  r = spreadsheet.eu_o19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0174,r[:number])
end

def test_eu_p19
  r = spreadsheet.eu_p19
  assert_equal(:ExcelString,r[:type])
  assert_equal("Note: 1720 \"based on current scope\"",r[:string].force_encoding('utf-8'))
end

def test_eu_b20
  r = spreadsheet.eu_b20
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total free allocation, MtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_f20
  r = spreadsheet.eu_f20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(848.7075677419355,r[:number])
end

def test_eu_g20
  r = spreadsheet.eu_g20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(763.1506248366406,r[:number])
end

def test_eu_h20
  r = spreadsheet.eu_h20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(677.1799459732,r[:number])
end

def test_eu_i20
  r = spreadsheet.eu_i20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(590.7938795210337,r[:number])
end

def test_eu_j20
  r = spreadsheet.eu_j20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(503.9907679979854,r[:number])
end

def test_eu_k20
  r = spreadsheet.eu_k20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(416.7689480508982,r[:number])
end

def test_eu_l20
  r = spreadsheet.eu_l20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(329.1267504361288,r[:number])
end

def test_eu_m20
  r = spreadsheet.eu_m20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(241.06249999999994,r[:number])
end

def test_eu_b21
  r = spreadsheet.eu_b21
  assert_equal(:ExcelString,r[:type])
  assert_equal("Volume available for auctioning",r[:string].force_encoding('utf-8'))
end

def test_eu_f21
  r = spreadsheet.eu_f21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1243.0870322580647,r[:number])
end

def test_eu_g21
  r = spreadsheet.eu_g21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1292.2467491233595,r[:number])
end

def test_eu_h21
  r = spreadsheet.eu_h21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1342.4535136798963,r[:number])
end

def test_eu_i21
  r = spreadsheet.eu_i21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1393.6979579340987,r[:number])
end

def test_eu_j21
  r = spreadsheet.eu_j21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1445.9709114854277,r[:number])
end

def test_eu_k21
  r = spreadsheet.eu_k21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1499.2633982095035,r[:number])
end

def test_eu_l21
  r = spreadsheet.eu_l21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1553.5666329993421,r[:number])
end

def test_eu_m21
  r = spreadsheet.eu_m21
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1608.8720185636937,r[:number])
end

def test_eu_b23
  r = spreadsheet.eu_b23
  assert_equal(:ExcelString,r[:type])
  assert_equal("Carbon Price",r[:string].force_encoding('utf-8'))
end

def test_eu_f23
  r = spreadsheet.eu_f23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g23
  r = spreadsheet.eu_g23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h23
  r = spreadsheet.eu_h23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i23
  r = spreadsheet.eu_i23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j23
  r = spreadsheet.eu_j23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k23
  r = spreadsheet.eu_k23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l23
  r = spreadsheet.eu_l23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m23
  r = spreadsheet.eu_m23
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_b24
  r = spreadsheet.eu_b24
  assert_equal(:ExcelString,r[:type])
  assert_equal("Price per allowance",r[:string].force_encoding('utf-8'))
end

def test_eu_f24
  r = spreadsheet.eu_f24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(25.0,r[:number])
end

def test_eu_g24
  r = spreadsheet.eu_g24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(26.25,r[:number])
end

def test_eu_h24
  r = spreadsheet.eu_h24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(27.5625,r[:number])
end

def test_eu_i24
  r = spreadsheet.eu_i24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(28.940625,r[:number])
end

def test_eu_j24
  r = spreadsheet.eu_j24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(30.387656250000003,r[:number])
end

def test_eu_k24
  r = spreadsheet.eu_k24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(31.907039062500004,r[:number])
end

def test_eu_l24
  r = spreadsheet.eu_l24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(33.50239101562501,r[:number])
end

def test_eu_m24
  r = spreadsheet.eu_m24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(35.17751056640626,r[:number])
end

def test_eu_n24
  r = spreadsheet.eu_n24
  assert_equal(:ExcelString,r[:type])
  assert_equal("€/tCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_o24
  r = spreadsheet.eu_o24
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.05,r[:number])
end

def test_eu_b25
  r = spreadsheet.eu_b25
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total revenue from auctions",r[:string].force_encoding('utf-8'))
end

def test_eu_f25
  r = spreadsheet.eu_f25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(31077.175806451618,r[:number])
end

def test_eu_g25
  r = spreadsheet.eu_g25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(33921.477164488184,r[:number])
end

def test_eu_h25
  r = spreadsheet.eu_h25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(37001.37497080214,r[:number])
end

def test_eu_i25
  r = spreadsheet.eu_i25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(40334.489963836524,r[:number])
end

def test_eu_j25
  r = spreadsheet.eu_j25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(43939.66700571836,r[:number])
end

def test_eu_k25
  r = spreadsheet.eu_k25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(47837.05581164713,r[:number])
end

def test_eu_l25
  r = spreadsheet.eu_l25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(52048.196807571956,r[:number])
end

def test_eu_m25
  r = spreadsheet.eu_m25
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(56596.112433019705,r[:number])
end

def test_eu_b27
  r = spreadsheet.eu_b27
  assert_equal(:ExcelString,r[:type])
  assert_equal("EU-27 Auction volumes- bought into proportion to net shortfall",r[:string].force_encoding('utf-8'))
end

def test_eu_f27
  r = spreadsheet.eu_f27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g27
  r = spreadsheet.eu_g27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h27
  r = spreadsheet.eu_h27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i27
  r = spreadsheet.eu_i27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j27
  r = spreadsheet.eu_j27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k27
  r = spreadsheet.eu_k27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l27
  r = spreadsheet.eu_l27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m27
  r = spreadsheet.eu_m27
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_b28
  r = spreadsheet.eu_b28
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_eu_f28
  r = spreadsheet.eu_f28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1084.4557258064517,r[:number])
end

def test_eu_g28
  r = spreadsheet.eu_g28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1065.5861961774194,r[:number])
end

def test_eu_h28
  r = spreadsheet.eu_h28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1047.0449963639326,r[:number])
end

def test_eu_i28
  r = spreadsheet.eu_i28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1028.8264134272,r[:number])
end

def test_eu_j28
  r = spreadsheet.eu_j28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1010.9248338335667,r[:number])
end

def test_eu_k28
  r = spreadsheet.eu_k28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(993.3347417248626,r[:number])
end

def test_eu_l28
  r = spreadsheet.eu_l28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(976.05071721885,r[:number])
end

def test_eu_m28
  r = spreadsheet.eu_m28
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(959.0674347392421,r[:number])
end

def test_eu_b29
  r = spreadsheet.eu_b29
  assert_equal(:ExcelString,r[:type])
  assert_equal("Leakage sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_f29
  r = spreadsheet.eu_f29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(32.0516914516129,r[:number])
end

def test_eu_g29
  r = spreadsheet.eu_g29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(39.367490025443544,r[:number])
end

def test_eu_h29
  r = spreadsheet.eu_h29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(46.41899483880101,r[:number])
end

def test_eu_i29
  r = spreadsheet.eu_i29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(53.21318838337351,r[:number])
end

def test_eu_j29
  r = spreadsheet.eu_j29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(59.756890177717516,r[:number])
end

def test_eu_k29
  r = spreadsheet.eu_k29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(66.05676032470339,r[:number])
end

def test_eu_l29
  r = spreadsheet.eu_l29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(72.11930299450394,r[:number])
end

def test_eu_m29
  r = spreadsheet.eu_m29
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(77.95086983463953,r[:number])
end

def test_eu_b30
  r = spreadsheet.eu_b30
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other",r[:string].force_encoding('utf-8'))
end

def test_eu_f30
  r = spreadsheet.eu_f30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(137.36439193548387,r[:number])
end

def test_eu_g30
  r = spreadsheet.eu_g30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(212.10239523912443,r[:number])
end

def test_eu_h30
  r = spreadsheet.eu_h30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(284.197927584496,r[:number])
end

def test_eu_i30
  r = spreadsheet.eu_i30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(353.72031928306586,r[:number])
end

def test_eu_j30
  r = spreadsheet.eu_j30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(420.7372879859702,r[:number])
end

def test_eu_k30
  r = spreadsheet.eu_k30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(485.31497381414727,r[:number])
end

def test_eu_l30
  r = spreadsheet.eu_l30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(547.5179737541932,r[:number])
end

def test_eu_m30
  r = spreadsheet.eu_m30
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(607.4093753348534,r[:number])
end

def test_eu_b31
  r = spreadsheet.eu_b31
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_eu_f31
  r = spreadsheet.eu_f31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1253.8718091935484,r[:number])
end

def test_eu_g31
  r = spreadsheet.eu_g31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1317.0560814419875,r[:number])
end

def test_eu_h31
  r = spreadsheet.eu_h31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1377.6619187872295,r[:number])
end

def test_eu_i31
  r = spreadsheet.eu_i31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1435.7599210936394,r[:number])
end

def test_eu_j31
  r = spreadsheet.eu_j31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1491.4190119972543,r[:number])
end

def test_eu_k31
  r = spreadsheet.eu_k31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1544.7064758637134,r[:number])
end

def test_eu_l31
  r = spreadsheet.eu_l31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1595.687993967547,r[:number])
end

def test_eu_m31
  r = spreadsheet.eu_m31
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1644.427679908735,r[:number])
end

def test_eu_b33
  r = spreadsheet.eu_b33
  assert_equal(:ExcelString,r[:type])
  assert_equal("Revenues",r[:string].force_encoding('utf-8'))
end

def test_eu_b34
  r = spreadsheet.eu_b34
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_eu_f34
  r = spreadsheet.eu_f34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(27111.39314516129,r[:number])
end

def test_eu_g34
  r = spreadsheet.eu_g34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(27971.63764965726,r[:number])
end

def test_eu_h34
  r = spreadsheet.eu_h34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(28859.177712280893,r[:number])
end

def test_eu_i34
  r = spreadsheet.eu_i34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(29774.87942109156,r[:number])
end

def test_eu_j34
  r = spreadsheet.eu_j34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(30719.6363451228,r[:number])
end

def test_eu_k34
  r = spreadsheet.eu_k34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(31694.370406353544,r[:number])
end

def test_eu_l34
  r = spreadsheet.eu_l34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(32700.032779347144,r[:number])
end

def test_eu_m34
  r = spreadsheet.eu_m34
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(33737.60481943584,r[:number])
end

def test_eu_b35
  r = spreadsheet.eu_b35
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other sectors",r[:string].force_encoding('utf-8'))
end

def test_eu_f35
  r = spreadsheet.eu_f35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4235.402084677419,r[:number])
end

def test_eu_g35
  r = spreadsheet.eu_g35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(6601.08448819491,r[:number])
end

def test_eu_h35
  r = spreadsheet.eu_h35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(9112.628924292123,r[:number])
end

def test_eu_i35
  r = spreadsheet.eu_i35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(11776.910045309047,r[:number])
end

def test_eu_j35
  r = spreadsheet.eu_j35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(14601.091916164401,r[:number])
end

def test_eu_k35
  r = spreadsheet.eu_k35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(17592.639459126676,r[:number])
end

def test_eu_l35
  r = spreadsheet.eu_l35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20759.3303334919,r[:number])
end

def test_eu_m35
  r = spreadsheet.eu_m35
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(24109.267266244624,r[:number])
end

def test_eu_b39
  r = spreadsheet.eu_b39
  assert_equal(:ExcelString,r[:type])
  assert_equal("2005 UK ETS Emissions",r[:string].force_encoding('utf-8'))
end

def test_eu_f39
  r = spreadsheet.eu_f39
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(242.0,r[:number])
end

def test_eu_g39
  r = spreadsheet.eu_g39
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_n39
  r = spreadsheet.eu_n39
  assert_equal(:ExcelString,r[:type])
  assert_equal(" ",r[:string].force_encoding('utf-8'))
end

def test_eu_b40
  r = spreadsheet.eu_b40
  assert_equal(:ExcelString,r[:type])
  assert_equal("2005 EU ETS Emissions",r[:string].force_encoding('utf-8'))
end

def test_eu_f40
  r = spreadsheet.eu_f40
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1785.0,r[:number])
end

def test_eu_g40
  r = spreadsheet.eu_g40
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_eu_b41
  r = spreadsheet.eu_b41
  assert_equal(:ExcelString,r[:type])
  assert_equal("Basic UK share of auction revenues",r[:string].force_encoding('utf-8'))
end

def test_eu_f41
  r = spreadsheet.eu_f41
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.13557422969187674,r[:number])
end

def test_eu_b42
  r = spreadsheet.eu_b42
  assert_equal(:ExcelString,r[:type])
  assert_equal("Amount of share auctioned in UK",r[:string].force_encoding('utf-8'))
end

def test_eu_f42
  r = spreadsheet.eu_f42
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.9,r[:number])
end

def test_eu_b43
  r = spreadsheet.eu_b43
  assert_equal(:ExcelString,r[:type])
  assert_equal("Actual UK share of auction revenues",r[:string].force_encoding('utf-8'))
end

def test_eu_f43
  r = spreadsheet.eu_f43
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.12201680672268907,r[:number])
end

def test_eu_b45
  r = spreadsheet.eu_b45
  assert_equal(:ExcelString,r[:type])
  assert_equal("UK Auction revenues",r[:string].force_encoding('utf-8'))
end

def test_eu_f45
  r = spreadsheet.eu_f45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_eu_g45
  r = spreadsheet.eu_g45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_eu_h45
  r = spreadsheet.eu_h45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_eu_i45
  r = spreadsheet.eu_i45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_eu_j45
  r = spreadsheet.eu_j45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_eu_k45
  r = spreadsheet.eu_k45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_eu_l45
  r = spreadsheet.eu_l45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_eu_m45
  r = spreadsheet.eu_m45
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_eu_b46
  r = spreadsheet.eu_b46
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_eu_f46
  r = spreadsheet.eu_f46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(152.99343419739765,r[:number])
end

def test_eu_g46
  r = spreadsheet.eu_g46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(160.70297733224922,r[:number])
end

def test_eu_h46
  r = spreadsheet.eu_h46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(168.09790807387034,r[:number])
end

def test_eu_i46
  r = spreadsheet.eu_i46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(175.1868407922659,r[:number])
end

def test_eu_j46
  r = spreadsheet.eu_j46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(181.97818532941287,r[:number])
end

def test_eu_k46
  r = spreadsheet.eu_k46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(188.4801515087489,r[:number])
end

def test_eu_l46
  r = spreadsheet.eu_l46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(194.70075354965363,r[:number])
end

def test_eu_m46
  r = spreadsheet.eu_m46
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(200.6478143888641,r[:number])
end

def test_eu_n46
  r = spreadsheet.eu_n46
  assert_equal(:ExcelString,r[:type])
  assert_equal("€bn",r[:string].force_encoding('utf-8'))
end


  # start of OLD UK
def test_old_uk_b2
  r = spreadsheet.old_uk_b2
  assert_equal(:ExcelString,r[:type])
  assert_equal("UK - Possible phase III auction revenues",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b4
  r = spreadsheet.old_uk_b4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Emissions",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c4
  r = spreadsheet.old_uk_c4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_old_uk_d4
  r = spreadsheet.old_uk_d4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_old_uk_e4
  r = spreadsheet.old_uk_e4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_old_uk_f4
  r = spreadsheet.old_uk_f4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_old_uk_g4
  r = spreadsheet.old_uk_g4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_old_uk_h4
  r = spreadsheet.old_uk_h4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_old_uk_i4
  r = spreadsheet.old_uk_i4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_old_uk_j4
  r = spreadsheet.old_uk_j4
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_old_uk_l4
  r = spreadsheet.old_uk_l4
  assert_equal(:ExcelString,r[:type])
  assert_equal("Change",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b5
  r = spreadsheet.old_uk_b5
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c5
  r = spreadsheet.old_uk_c5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(163.25,r[:number])
end

def test_old_uk_d5
  r = spreadsheet.old_uk_d5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(160.40945,r[:number])
end

def test_old_uk_e5
  r = spreadsheet.old_uk_e5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(157.61832557,r[:number])
end

def test_old_uk_f5
  r = spreadsheet.old_uk_f5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(154.87576670508201,r[:number])
end

def test_old_uk_g5
  r = spreadsheet.old_uk_g5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(152.1809283644136,r[:number])
end

def test_old_uk_h5
  r = spreadsheet.old_uk_h5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(149.5329802108728,r[:number])
end

def test_old_uk_i5
  r = spreadsheet.old_uk_i5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(146.93110635520364,r[:number])
end

def test_old_uk_j5
  r = spreadsheet.old_uk_j5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(144.3745051046231,r[:number])
end

def test_old_uk_k5
  r = spreadsheet.old_uk_k5
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO2",r[:string].force_encoding('utf-8'))
end

def test_old_uk_l5
  r = spreadsheet.old_uk_l5
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0174,r[:number])
end

def test_old_uk_b6
  r = spreadsheet.old_uk_b6
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c6
  r = spreadsheet.old_uk_c6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(134.16697500000004,r[:number])
end

def test_old_uk_d6
  r = spreadsheet.old_uk_d6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(131.83246963500005,r[:number])
end

def test_old_uk_e6
  r = spreadsheet.old_uk_e6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(129.53858466335103,r[:number])
end

def test_old_uk_f6
  r = spreadsheet.old_uk_f6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(127.28461329020871,r[:number])
end

def test_old_uk_g6
  r = spreadsheet.old_uk_g6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(125.06986101895905,r[:number])
end

def test_old_uk_h6
  r = spreadsheet.old_uk_h6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(122.89364543722914,r[:number])
end

def test_old_uk_i6
  r = spreadsheet.old_uk_i6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(120.75529600662136,r[:number])
end

def test_old_uk_j6
  r = spreadsheet.old_uk_j6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(118.65415385610618,r[:number])
end

def test_old_uk_k6
  r = spreadsheet.old_uk_k6
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO3",r[:string].force_encoding('utf-8'))
end

def test_old_uk_l6
  r = spreadsheet.old_uk_l6
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.017400000000000082,r[:number])
end

def test_old_uk_b7
  r = spreadsheet.old_uk_b7
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c7
  r = spreadsheet.old_uk_c7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(297.41697500000004,r[:number])
end

def test_old_uk_d7
  r = spreadsheet.old_uk_d7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(292.24191963500004,r[:number])
end

def test_old_uk_e7
  r = spreadsheet.old_uk_e7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(287.156910233351,r[:number])
end

def test_old_uk_f7
  r = spreadsheet.old_uk_f7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(282.1603799952907,r[:number])
end

def test_old_uk_g7
  r = spreadsheet.old_uk_g7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(277.25078938337265,r[:number])
end

def test_old_uk_h7
  r = spreadsheet.old_uk_h7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(272.42662564810195,r[:number])
end

def test_old_uk_i7
  r = spreadsheet.old_uk_i7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(267.686402361825,r[:number])
end

def test_old_uk_j7
  r = spreadsheet.old_uk_j7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(263.0286589607293,r[:number])
end

def test_old_uk_k7
  r = spreadsheet.old_uk_k7
  assert_equal(:ExcelString,r[:type])
  assert_equal("mtCO4",r[:string].force_encoding('utf-8'))
end

def test_old_uk_l7
  r = spreadsheet.old_uk_l7
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.0174,r[:number])
end

def test_old_uk_b9
  r = spreadsheet.old_uk_b9
  assert_equal(:ExcelString,r[:type])
  assert_equal("Proportion auctioned",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c9
  r = spreadsheet.old_uk_c9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_old_uk_d9
  r = spreadsheet.old_uk_d9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_old_uk_e9
  r = spreadsheet.old_uk_e9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_old_uk_f9
  r = spreadsheet.old_uk_f9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_old_uk_g9
  r = spreadsheet.old_uk_g9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_old_uk_h9
  r = spreadsheet.old_uk_h9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_old_uk_i9
  r = spreadsheet.old_uk_i9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_old_uk_j9
  r = spreadsheet.old_uk_j9
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_old_uk_b10
  r = spreadsheet.old_uk_b10
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c10
  r = spreadsheet.old_uk_c10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_d10
  r = spreadsheet.old_uk_d10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_e10
  r = spreadsheet.old_uk_e10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_f10
  r = spreadsheet.old_uk_f10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_g10
  r = spreadsheet.old_uk_g10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_h10
  r = spreadsheet.old_uk_h10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_i10
  r = spreadsheet.old_uk_i10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_j10
  r = spreadsheet.old_uk_j10
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_k10
  r = spreadsheet.old_uk_k10
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b11
  r = spreadsheet.old_uk_b11
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c11
  r = spreadsheet.old_uk_c11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.2,r[:number])
end

def test_old_uk_d11
  r = spreadsheet.old_uk_d11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.3142857142857143,r[:number])
end

def test_old_uk_e11
  r = spreadsheet.old_uk_e11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.4285714285714286,r[:number])
end

def test_old_uk_f11
  r = spreadsheet.old_uk_f11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5428571428571429,r[:number])
end

def test_old_uk_g11
  r = spreadsheet.old_uk_g11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.6571428571428573,r[:number])
end

def test_old_uk_h11
  r = spreadsheet.old_uk_h11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7714285714285716,r[:number])
end

def test_old_uk_i11
  r = spreadsheet.old_uk_i11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8857142857142859,r[:number])
end

def test_old_uk_j11
  r = spreadsheet.old_uk_j11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_k11
  r = spreadsheet.old_uk_k11
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_old_uk_l11
  r = spreadsheet.old_uk_l11
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.1142857142857143,r[:number])
end

def test_old_uk_b12
  r = spreadsheet.old_uk_b12
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c12
  r = spreadsheet.old_uk_c12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.6391141426947805,r[:number])
end

def test_old_uk_d12
  r = spreadsheet.old_uk_d12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.6906692651669547,r[:number])
end

def test_old_uk_e12
  r = spreadsheet.old_uk_e12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.742224387639129,r[:number])
end

def test_old_uk_f12
  r = spreadsheet.old_uk_f12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.7937795101113032,r[:number])
end

def test_old_uk_g12
  r = spreadsheet.old_uk_g12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8453346325834775,r[:number])
end

def test_old_uk_h12
  r = spreadsheet.old_uk_h12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8968897550556517,r[:number])
end

def test_old_uk_i12
  r = spreadsheet.old_uk_i12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.9484448775278258,r[:number])
end

def test_old_uk_j12
  r = spreadsheet.old_uk_j12
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.0,r[:number])
end

def test_old_uk_k12
  r = spreadsheet.old_uk_k12
  assert_equal(:ExcelString,r[:type])
  assert_equal("%",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c14
  r = spreadsheet.old_uk_c14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_old_uk_d14
  r = spreadsheet.old_uk_d14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_old_uk_e14
  r = spreadsheet.old_uk_e14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_old_uk_f14
  r = spreadsheet.old_uk_f14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_old_uk_g14
  r = spreadsheet.old_uk_g14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_old_uk_h14
  r = spreadsheet.old_uk_h14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_old_uk_i14
  r = spreadsheet.old_uk_i14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_old_uk_j14
  r = spreadsheet.old_uk_j14
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_old_uk_b15
  r = spreadsheet.old_uk_b15
  assert_equal(:ExcelString,r[:type])
  assert_equal("CO2 price",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c15
  r = spreadsheet.old_uk_c15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_d15
  r = spreadsheet.old_uk_d15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_e15
  r = spreadsheet.old_uk_e15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_f15
  r = spreadsheet.old_uk_f15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_g15
  r = spreadsheet.old_uk_g15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_h15
  r = spreadsheet.old_uk_h15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_i15
  r = spreadsheet.old_uk_i15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_j15
  r = spreadsheet.old_uk_j15
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(20.0,r[:number])
end

def test_old_uk_k15
  r = spreadsheet.old_uk_k15
  assert_equal(:ExcelString,r[:type])
  assert_equal("€/tCO2",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b17
  r = spreadsheet.old_uk_b17
  assert_equal(:ExcelString,r[:type])
  assert_equal("Auction revenues",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c17
  r = spreadsheet.old_uk_c17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2013.0,r[:number])
end

def test_old_uk_d17
  r = spreadsheet.old_uk_d17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2014.0,r[:number])
end

def test_old_uk_e17
  r = spreadsheet.old_uk_e17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2015.0,r[:number])
end

def test_old_uk_f17
  r = spreadsheet.old_uk_f17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2016.0,r[:number])
end

def test_old_uk_g17
  r = spreadsheet.old_uk_g17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2017.0,r[:number])
end

def test_old_uk_h17
  r = spreadsheet.old_uk_h17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2018.0,r[:number])
end

def test_old_uk_i17
  r = spreadsheet.old_uk_i17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2019.0,r[:number])
end

def test_old_uk_j17
  r = spreadsheet.old_uk_j17
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2020.0,r[:number])
end

def test_old_uk_b18
  r = spreadsheet.old_uk_b18
  assert_equal(:ExcelString,r[:type])
  assert_equal("Power sector",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c18
  r = spreadsheet.old_uk_c18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.265,r[:number])
end

def test_old_uk_d18
  r = spreadsheet.old_uk_d18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.208189,r[:number])
end

def test_old_uk_e18
  r = spreadsheet.old_uk_e18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.1523665114,r[:number])
end

def test_old_uk_f18
  r = spreadsheet.old_uk_f18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.09751533410164,r[:number])
end

def test_old_uk_g18
  r = spreadsheet.old_uk_g18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.043618567288272,r[:number])
end

def test_old_uk_h18
  r = spreadsheet.old_uk_h18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2.990659604217456,r[:number])
end

def test_old_uk_i18
  r = spreadsheet.old_uk_i18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2.9386221271040727,r[:number])
end

def test_old_uk_j18
  r = spreadsheet.old_uk_j18
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2.8874901020924617,r[:number])
end

def test_old_uk_k18
  r = spreadsheet.old_uk_k18
  assert_equal(:ExcelString,r[:type])
  assert_equal("€bn",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b19
  r = spreadsheet.old_uk_b19
  assert_equal(:ExcelString,r[:type])
  assert_equal("Other",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c19
  r = spreadsheet.old_uk_c19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.5366679000000002,r[:number])
end

def test_old_uk_d19
  r = spreadsheet.old_uk_d19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(0.8286612377057146,r[:number])
end

def test_old_uk_e19
  r = spreadsheet.old_uk_e19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.110330725685866,r[:number])
end

def test_old_uk_f19
  r = spreadsheet.old_uk_f19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.3819472300079803,r[:number])
end

def test_old_uk_g19
  r = spreadsheet.old_uk_g19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.6437753162491764,r[:number])
end

def test_old_uk_h19
  r = spreadsheet.old_uk_h19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(1.8960733867458215,r[:number])
end

def test_old_uk_i19
  r = spreadsheet.old_uk_i19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2.1390938149744363,r[:number])
end

def test_old_uk_j19
  r = spreadsheet.old_uk_j19
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(2.3730830771221236,r[:number])
end

def test_old_uk_k19
  r = spreadsheet.old_uk_k19
  assert_equal(:ExcelString,r[:type])
  assert_equal("€bn",r[:string].force_encoding('utf-8'))
end

def test_old_uk_b20
  r = spreadsheet.old_uk_b20
  assert_equal(:ExcelString,r[:type])
  assert_equal("Total",r[:string].force_encoding('utf-8'))
end

def test_old_uk_c20
  r = spreadsheet.old_uk_c20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(3.8016679,r[:number])
end

def test_old_uk_d20
  r = spreadsheet.old_uk_d20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4.036850237705714,r[:number])
end

def test_old_uk_e20
  r = spreadsheet.old_uk_e20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4.262697237085867,r[:number])
end

def test_old_uk_f20
  r = spreadsheet.old_uk_f20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4.479462564109621,r[:number])
end

def test_old_uk_g20
  r = spreadsheet.old_uk_g20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4.687393883537449,r[:number])
end

def test_old_uk_h20
  r = spreadsheet.old_uk_h20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(4.886732990963277,r[:number])
end

def test_old_uk_i20
  r = spreadsheet.old_uk_i20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(5.077715942078508,r[:number])
end

def test_old_uk_j20
  r = spreadsheet.old_uk_j20
  assert_equal(:ExcelNumber,r[:type])
  assert_in_epsilon(5.260573179214585,r[:number])
end

def test_old_uk_k20
  r = spreadsheet.old_uk_k20
  assert_equal(:ExcelString,r[:type])
  assert_equal("€bn",r[:string].force_encoding('utf-8'))
end

def test_old_uk_k22
  r = spreadsheet.old_uk_k22
  assert_equal(:ExcelString,r[:type])
  assert_equal(" ",r[:string].force_encoding('utf-8'))
end

end
