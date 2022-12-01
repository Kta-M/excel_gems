require_relative '../../../utils/memory'

# メモリ使用量
print_memory_usage do
  require 'rubyXL'
  workbook = RubyXL::Parser.parse('input/benchmark.xlsx')
  workbook[0].each do |row|
    # なにかする
  end
end
