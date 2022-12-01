require 'roo'

# Workbookを読み込み
workbook = Roo::Excelx.new('input/example.xlsx')

# Sheet取得
sheet = workbook.sheet('curry')

# 値の取得
sheet.first_row.upto(sheet.last_row) do |row|
  puts "#{sheet.cell(row, 1)} #{sheet.cell(row, 2)} #{sheet.cell(row, 3)} #{sheet.cell(row, 4)}"
end
