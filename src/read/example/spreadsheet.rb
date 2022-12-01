require 'spreadsheet'

# Workbookを読み込み
workbook = Spreadsheet.open('input/example.xls')

# Sheet取得
sheet = workbook.worksheet('curry')

# 値を取得
sheet.each do |row|
  # 数式が入っているセルは .value で値を取得
  puts "#{row[0]} #{row[1]} #{row[2]} #{row[3].instance_of?(Spreadsheet::Formula) ? row[3].value : row[3]}"
end
