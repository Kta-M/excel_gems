require_relative '../../../utils/memory'

# メモリ使用量
print_memory_usage do
  require 'roo'
  workbook = Roo::Excelx.new('input/benchmark.xlsx')
  workbook.each_row_streaming do |row|
    # なにかする
  end
end
