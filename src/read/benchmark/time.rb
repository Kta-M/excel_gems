require 'benchmark'
require 'roo'
require 'spreadsheet'
require 'rubyXL'
require 'rubyXL/convenience_methods'

Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  x.report("roo") do
    workbook = Roo::Excelx.new('input/benchmark.xlsx')
    workbook.each_row_streaming do |row|
      # なにかする
    end
  end

  x.report("spreadsheet") do
    workbook = Spreadsheet.open('input/benchmark.xls')
    workbook.worksheet(0).each do |row|
      # なにかする
    end
  end

  x.report("rubyxl") do
    workbook = RubyXL::Parser.parse('input/benchmark.xlsx')
    workbook[0].each do |row|
      # なにかする
    end
  end
end
