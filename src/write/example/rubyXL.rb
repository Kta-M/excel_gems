require 'rubyXL'
require 'rubyXL/convenience_methods'

# Workbookを作成
workbook = RubyXL::Workbook.new

# Sheet取得
sheet = workbook[0]   # 最初からシートが１つある

# Sheetの名前変更
sheet.sheet_name = 'curry'

# 各セルに書き込み
sheet.add_cell(0, 0, '品名')
sheet.add_cell(0, 1, '単価')
sheet.add_cell(0, 2, '数量')
sheet.add_cell(0, 3, '計')

sheet.add_cell(1, 0, 'にんじん')
sheet.add_cell(1, 1, 80)
sheet.add_cell(1, 2, 1)
sheet.add_cell(1, 3, '', 'B2*C2')

sheet.add_cell(2, 0, 'たまねぎ')
sheet.add_cell(2, 1, 50)
sheet.add_cell(2, 2, 2)
sheet.add_cell(2, 3, '', 'B3*C3')

sheet.add_cell(3, 0, 'じゃがいも')
sheet.add_cell(3, 1, 40)
sheet.add_cell(3, 2, 2)
sheet.add_cell(3, 3, '', 'B4*C4')

sheet.add_cell(4, 0, '牛肉')
sheet.add_cell(4, 1, 200)
sheet.add_cell(4, 2, 1)
sheet.add_cell(4, 3, '', 'B5*C5')

sheet.add_cell(5, 0, 'カレー粉')
sheet.add_cell(5, 1, 150)
sheet.add_cell(5, 2, 1)
sheet.add_cell(5, 3, '', 'B6*C6')

sheet.add_cell(6, 0, '総計')
sheet.add_cell(6, 1, '') # 値を入れないと罫線が引けない
sheet.add_cell(6, 2, '') # 値を入れないと罫線が引けない
sheet.add_cell(6, 3, '', 'SUM(D2:D6)')

# セルのマージ
sheet.merge_cells(6, 0, 6, 2)

# 各行のフォントを変更
(0..6).each do |row|
  sheet.change_row_font_name(row, 'メイリオ')
end

# 背景色設定
(0..3).each do |col|
  sheet[0][col].change_fill('d0d0d0')
end
sheet[6][3].change_fill('b8cce4')

# フォントカラー設定
sheet[6][3].change_font_color('ae2f29')

# 書体変更
sheet[6][3].change_font_bold(true)

# アラインメント
sheet.change_row_horizontal_alignment(0, 'center')
sheet[6][0].change_horizontal_alignment('right')

# 罫線
(0..6).each do |row|
  (0..3).each do |col|
    sheet[row][col].change_border(:top, 'thin')
    sheet[row][col].change_border(:left, 'thin')
    sheet[row][col].change_border(:bottom, 'thin')
    sheet[row][col].change_border(:right, 'thin')
  end
end

# 書き出し
workbook.write("output/example/rubyXL.xlsx")
