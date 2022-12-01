require_relative '../../../utils/memory'

SHEET_NUM = 100
ROW_NUM = 100
COL_NUM = 100

#==============================================================================

# メモリ使用量
print_memory_usage do
  require_relative '../../../utils/spreadsheet'
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  (0...ROW_NUM).each do |row|
    (0...COL_NUM).each do |col|
      sheet.row(row).update_format(col, font: Spreadsheet::Font.new('メイリオ'))
      sheet.row(row).update_format(col, pattern_fg_color: :silver, pattern: 1)
      sheet.row(row).update_format(col, color: :silver)
      sheet.row(row).update_format(col, weight: :bold)
      sheet.row(row).update_format(col, align: :center)
      sheet.row(row).update_format(col, border: :thin)
    end
  end

  sheet.workbook.write('output/benchmark/spreadsheet.xls')
end
