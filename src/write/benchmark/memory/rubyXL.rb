require_relative '../../../utils/memory'

SHEET_NUM = 100
ROW_NUM = 100
COL_NUM = 100

#==============================================================================

# メモリ使用量
print_memory_usage do
  require_relative '../../../utils/rubyXL'
  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  (0...ROW_NUM).each do |row|
    (0...COL_NUM).each do |col|
      sheet[row][col].change_font_name('メイリオ')
      sheet[row][col].change_fill('d0d0d0')
      sheet[row][col].change_font_color('d0d0d0')
      sheet[row][col].change_font_bold(true)
      sheet[row][col].change_horizontal_alignment('center')
      sheet[row][col].change_border(:top, 'thin')
      sheet[row][col].change_border(:left, 'thin')
      sheet[row][col].change_border(:bottom, 'thin')
      sheet[row][col].change_border(:right, 'thin')
    end
  end

  sheet.workbook.write('output/benchmark/rubyXL.xlsx')
end
