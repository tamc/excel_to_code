# coding: utf-8
# All tests for /Users/tamc/Documents/github/excel_to_code/examples/global.xlsx
require 'test/unit'
require_relative 'global'

class TestGlobal < Test::Unit::TestCase
  def worksheet; @worksheet ||= Global.new; end
  def test_g_30_b2; assert_equal("G", worksheet.g_30_b2); end
  def test_g_30_c2; assert_equal(:ref, worksheet.g_30_c2); end
  def test_g_30_e2; assert_equal("Click here to return to the \"contents\" page", worksheet.g_30_e2); end
  def test_g_30_b3; assert_in_epsilon(30.0, worksheet.g_30_b3, 0.002); end
  def test_g_30_c3; assert_equal(:ref, worksheet.g_30_c3); end
  def test_g_30_aa3; assert_equal("Supporting notes this way…", worksheet.g_30_aa3); end
  def test_g_30_b5; assert_equal("Energy calculations / results", worksheet.g_30_b5); end
  def test_g_30_s5; assert_equal("Notes on the calculations", worksheet.g_30_s5); end
  def test_g_30_c9; assert_equal("Code", worksheet.g_30_c9); end
  def test_g_30_d9; assert_equal("Tech description", worksheet.g_30_d9); end
  def test_g_30_e9; assert_equal("Notes", worksheet.g_30_e9); end
  def test_g_30_f9; assert_equal("Energy vector", worksheet.g_30_f9); end
  def test_g_30_g9; assert_equal("Energy description", worksheet.g_30_g9); end
  def test_g_30_h9; assert_equal("Unit", worksheet.g_30_h9); end
  def test_g_30_i9; assert_equal("2011", worksheet.g_30_i9); end
  def test_g_30_j9; assert_equal("2012", worksheet.g_30_j9); end
  def test_g_30_k9; assert_equal("2013", worksheet.g_30_k9); end
  def test_g_30_l9; assert_equal("2014", worksheet.g_30_l9); end
  def test_g_30_m9; assert_equal("2015", worksheet.g_30_m9); end
  def test_g_30_n9; assert_equal("2016", worksheet.g_30_n9); end
  def test_g_30_o9; assert_equal("2017", worksheet.g_30_o9); end
  def test_g_30_p9; assert_equal("2018", worksheet.g_30_p9); end
  def test_g_30_q9; assert_equal("2019", worksheet.g_30_q9); end
  def test_g_30_h10; assert_equal(:ref, worksheet.g_30_h10); end
  def test_g_30_c11; assert_equal("Steel.Oxygen.Electricity", worksheet.g_30_c11); end
  def test_g_30_d11; assert_equal(:ref, worksheet.g_30_d11); end
  def test_g_30_f11; assert_equal("G.E.01", worksheet.g_30_f11); end
  def test_g_30_g11; assert_equal(:ref, worksheet.g_30_g11); end
  def test_g_30_h11; assert_equal(:ref, worksheet.g_30_h11); end
  def test_g_30_i11; assert_in_epsilon(-2.308955743700514, worksheet.g_30_i11, 0.002); end
  def test_g_30_c12; assert_equal("Steel.Oxygen.Coal", worksheet.g_30_c12); end
  def test_g_30_d12; assert_equal(:ref, worksheet.g_30_d12); end
  def test_g_30_f12; assert_equal("G.FF.01", worksheet.g_30_f12); end
  def test_g_30_g12; assert_equal(:ref, worksheet.g_30_g12); end
  def test_g_30_h12; assert_equal(:ref, worksheet.g_30_h12); end
  def test_g_30_i12; assert_in_epsilon(-12.811296723305107, worksheet.g_30_i12, 0.002); end
  def test_g_30_c13; assert_equal("Steel.Oxygen.Oil", worksheet.g_30_c13); end
  def test_g_30_d13; assert_equal(:ref, worksheet.g_30_d13); end
  def test_g_30_f13; assert_equal("G.FF.02", worksheet.g_30_f13); end
  def test_g_30_g13; assert_equal(:ref, worksheet.g_30_g13); end
  def test_g_30_h13; assert_equal(:ref, worksheet.g_30_h13); end
  def test_g_30_i13; assert_in_delta(-0.6053795859240001, worksheet.g_30_i13, 0.002); end
  def test_g_30_c14; assert_equal("Steel.Oxygen.NaturalGas", worksheet.g_30_c14); end
  def test_g_30_d14; assert_equal(:ref, worksheet.g_30_d14); end
  def test_g_30_f14; assert_equal("G.FF.03", worksheet.g_30_f14); end
  def test_g_30_g14; assert_equal(:ref, worksheet.g_30_g14); end
  def test_g_30_h14; assert_equal(:ref, worksheet.g_30_h14); end
  def test_g_30_i14; assert_in_epsilon(-1.0517591114635856, worksheet.g_30_i14, 0.002); end
  def test_g_30_c15; assert_equal("Steel.Oxygen.SolidHydrocarbons", worksheet.g_30_c15); end
  def test_g_30_d15; assert_equal(:ref, worksheet.g_30_d15); end
  def test_g_30_f15; assert_equal("G.C.01", worksheet.g_30_f15); end
  def test_g_30_g15; assert_equal(:ref, worksheet.g_30_g15); end
  def test_g_30_h15; assert_equal(:ref, worksheet.g_30_h15); end
  def test_g_30_i15; assert_in_delta(-0.17590630860000006, worksheet.g_30_i15, 0.002); end
  def test_g_30_c16; assert_equal("Steel.Oxygen.Heat", worksheet.g_30_c16); end
  def test_g_30_d16; assert_equal(:ref, worksheet.g_30_d16); end
  def test_g_30_f16; assert_equal("G.H.01", worksheet.g_30_f16); end
  def test_g_30_g16; assert_equal(:ref, worksheet.g_30_g16); end
  def test_g_30_h16; assert_equal(:ref, worksheet.g_30_h16); end
  def test_g_30_i16; assert_in_delta(-0.803408652648, worksheet.g_30_i16, 0.002); end
  def test_g_30_c17; assert_equal("Steel.OxygenHisarna.Electricity", worksheet.g_30_c17); end
  def test_g_30_d17; assert_equal(:ref, worksheet.g_30_d17); end
  def test_g_30_f17; assert_equal("G.E.01", worksheet.g_30_f17); end
  def test_g_30_g17; assert_equal(:ref, worksheet.g_30_g17); end
  def test_g_30_h17; assert_equal(:ref, worksheet.g_30_h17); end
  def test_g_30_i17; assert_in_delta(0.0, (worksheet.g_30_i17||0), 0.002); end
  def test_g_30_c18; assert_equal("Steel.OxygenHisarna.Coal", worksheet.g_30_c18); end
  def test_g_30_d18; assert_equal(:ref, worksheet.g_30_d18); end
  def test_g_30_f18; assert_equal("G.FF.01", worksheet.g_30_f18); end
  def test_g_30_g18; assert_equal(:ref, worksheet.g_30_g18); end
  def test_g_30_h18; assert_equal(:ref, worksheet.g_30_h18); end
  def test_g_30_i18; assert_in_delta(0.0, (worksheet.g_30_i18||0), 0.002); end
  def test_g_30_c19; assert_equal("Steel.OxygenHisarna.Oil", worksheet.g_30_c19); end
  def test_g_30_d19; assert_equal(:ref, worksheet.g_30_d19); end
  def test_g_30_f19; assert_equal("G.FF.02", worksheet.g_30_f19); end
  def test_g_30_g19; assert_equal(:ref, worksheet.g_30_g19); end
  def test_g_30_h19; assert_equal(:ref, worksheet.g_30_h19); end
  def test_g_30_i19; assert_in_delta(0.0, (worksheet.g_30_i19||0), 0.002); end
  def test_g_30_c20; assert_equal("Steel.OxygenHisarna.NaturalGas", worksheet.g_30_c20); end
  def test_g_30_d20; assert_equal(:ref, worksheet.g_30_d20); end
  def test_g_30_f20; assert_equal("G.FF.03", worksheet.g_30_f20); end
  def test_g_30_g20; assert_equal(:ref, worksheet.g_30_g20); end
  def test_g_30_h20; assert_equal(:ref, worksheet.g_30_h20); end
  def test_g_30_i20; assert_in_delta(0.0, (worksheet.g_30_i20||0), 0.002); end
  def test_g_30_c21; assert_equal("Steel.OxygenHisarna.SolidHydrocarbons", worksheet.g_30_c21); end
  def test_g_30_d21; assert_equal(:ref, worksheet.g_30_d21); end
  def test_g_30_f21; assert_equal("G.C.01", worksheet.g_30_f21); end
  def test_g_30_g21; assert_equal(:ref, worksheet.g_30_g21); end
  def test_g_30_h21; assert_equal(:ref, worksheet.g_30_h21); end
  def test_g_30_i21; assert_in_delta(0.0, (worksheet.g_30_i21||0), 0.002); end
  def test_g_30_c22; assert_equal("Steel.OxygenHisarna.Heat", worksheet.g_30_c22); end
  def test_g_30_d22; assert_equal(:ref, worksheet.g_30_d22); end
  def test_g_30_f22; assert_equal("G.H.01", worksheet.g_30_f22); end
  def test_g_30_g22; assert_equal(:ref, worksheet.g_30_g22); end
  def test_g_30_h22; assert_equal(:ref, worksheet.g_30_h22); end
  def test_g_30_i22; assert_in_delta(0.0, (worksheet.g_30_i22||0), 0.002); end
  def test_g_30_c23; assert_equal("Steel.Electric.Electricity", worksheet.g_30_c23); end
  def test_g_30_d23; assert_equal(:ref, worksheet.g_30_d23); end
  def test_g_30_f23; assert_equal("G.E.01", worksheet.g_30_f23); end
  def test_g_30_g23; assert_equal(:ref, worksheet.g_30_g23); end
  def test_g_30_h23; assert_equal(:ref, worksheet.g_30_h23); end
  def test_g_30_i23; assert_in_epsilon(-1.6352704427314861, worksheet.g_30_i23, 0.002); end
  def test_g_30_c24; assert_equal("Steel.Electric.Coal", worksheet.g_30_c24); end
  def test_g_30_d24; assert_equal(:ref, worksheet.g_30_d24); end
  def test_g_30_f24; assert_equal("G.FF.01", worksheet.g_30_f24); end
  def test_g_30_g24; assert_equal(:ref, worksheet.g_30_g24); end
  def test_g_30_h24; assert_equal(:ref, worksheet.g_30_h24); end
  def test_g_30_i24; assert_in_epsilon(-10.424196405786889, worksheet.g_30_i24, 0.002); end
  def test_g_30_c25; assert_equal("Steel.Electric.Oil", worksheet.g_30_c25); end
  def test_g_30_d25; assert_equal(:ref, worksheet.g_30_d25); end
  def test_g_30_f25; assert_equal("G.FF.02", worksheet.g_30_f25); end
  def test_g_30_g25; assert_equal(:ref, worksheet.g_30_g25); end
  def test_g_30_h25; assert_equal(:ref, worksheet.g_30_h25); end
  def test_g_30_i25; assert_in_delta(0.0, (worksheet.g_30_i25||0), 0.002); end
  def test_g_30_c26; assert_equal("Steel.Electric.NaturalGas", worksheet.g_30_c26); end
  def test_g_30_d26; assert_equal(:ref, worksheet.g_30_d26); end
  def test_g_30_f26; assert_equal("G.FF.03", worksheet.g_30_f26); end
  def test_g_30_g26; assert_equal(:ref, worksheet.g_30_g26); end
  def test_g_30_h26; assert_equal(:ref, worksheet.g_30_h26); end
  def test_g_30_i26; assert_in_epsilon(-1.2528070032244134, worksheet.g_30_i26, 0.002); end
  def test_g_30_c27; assert_equal("Steel.Electric.SolidHydrocarbons", worksheet.g_30_c27); end
  def test_g_30_d27; assert_equal(:ref, worksheet.g_30_d27); end
  def test_g_30_f27; assert_equal("G.C.01", worksheet.g_30_f27); end
  def test_g_30_g27; assert_equal(:ref, worksheet.g_30_g27); end
  def test_g_30_h27; assert_equal(:ref, worksheet.g_30_h27); end
  def test_g_30_i27; assert_in_delta(0.0, (worksheet.g_30_i27||0), 0.002); end
  def test_g_30_c28; assert_equal("Steel.Electric.Heat", worksheet.g_30_c28); end
  def test_g_30_d28; assert_equal(:ref, worksheet.g_30_d28); end
  def test_g_30_f28; assert_equal("G.H.01", worksheet.g_30_f28); end
  def test_g_30_g28; assert_equal(:ref, worksheet.g_30_g28); end
  def test_g_30_h28; assert_equal(:ref, worksheet.g_30_h28); end
  def test_g_30_i28; assert_in_delta(0.0, (worksheet.g_30_i28||0), 0.002); end
end
