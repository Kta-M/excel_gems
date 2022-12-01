require 'rubyXL'
require 'rubyXL/convenience_methods'

# Workbookを読み込み
workbook = RubyXL::Parser.parse('input/example.xlsx')

# Sheet取得
sheet = workbook['curry']

# 値を取得
sheet.each do |row|
  puts "#{row[0].value} #{row[1].value} #{row[2].value} #{row[3].value}"
end
