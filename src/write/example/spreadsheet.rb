require 'spreadsheet'

# Workbookを作成
workbook = Spreadsheet::Workbook.new

# Sheetを作成
sheet = workbook.create_worksheet(name: 'curry')

# 各セルに書き込み
# いろいろな方法でデータを入れられる
# 計算式は入力できない
sheet.row(0).concat %w{品名 単価 数量 計}

sheet[1,0] = 'にんじん'
sheet[1,1] = 80
sheet[1,2] = 1
sheet[1,3] = 80

row = sheet.row(2)
row.push 50
row.push 2
row.push 100
row.unshift 'たまねぎ'

sheet.row(3).replace [ 'じゃがいも', 40, 2, 80 ]

sheet.row(4).push '牛肉', 1, 200
sheet.row(4).insert 1, 200

sheet.update_row 5, 'カレー粉', 150, 1, 150

sheet[6,0] = '総計'
sheet[6,3] = 610

# セルのマージ
sheet.merge_cells(6, 0, 6, 2)

# 各行のフォントを変更
format = Spreadsheet::Format.new(font: Spreadsheet::Font.new('メイリオ'))
(0..6).each do |row|
  sheet.row(row).default_format = format
end

# 背景色設定
# パレットからの選択(カラーコード指定不可)
# https://github.com/zdavatz/spreadsheet/blob/master/lib/spreadsheet/datatypes.rb#L15-L77
(0..3).each do |col|
  sheet.row(0).update_format(col, pattern_fg_color: :silver, pattern: 1)
end
sheet.row(6).update_format(3, pattern_fg_color: :xls_color_16, pattern: 1)

# フォントカラー設定
sheet.row(6).update_format(3, color: :red)

# 書体変更
sheet.row(6).update_format(3, weight: :bold)

# アラインメント
(0..3).each do |col|
  sheet.row(0).update_format(col, align: :center)
end
sheet.row(6).update_format(0, align: :right)

# 罫線
(0..6).each do |row|
  (0..3).each do |col|
    sheet.row(row).update_format(col, border: :thin)
  end
end

# 書き出し
workbook.write('output/example/spreadsheet.xls')
