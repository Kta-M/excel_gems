require_relative '../../../utils/memory'

# メモリ使用量
print_memory_usage do
  require 'spreadsheet'
  workbook = Spreadsheet.open('input/benchmark.xls')
  workbook.worksheet(0).each do |row|
    # なにかする
  end
end
