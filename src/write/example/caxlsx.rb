require 'caxlsx'

# Workbookを作成
package = Axlsx::Package.new
workbook = package.workbook

# Sheetを作成
sheet = workbook.add_worksheet(name: 'curry')

# 各セルに書き込み
sheet.add_row(['品名',        '単価', '数量', '計'])
sheet.add_row(['にんじん',    80,     1,      '=B2*C2'])
sheet.add_row(['たまねぎ',    50,     2,      '=B3*C3'])
sheet.add_row(['じゃがいも',  40,     2,      '=B4*C4'])
sheet.add_row(['牛肉',        200,    1,      '=B5*C5'])
sheet.add_row(['カレー粉',    150,    1,      '=B6*C6'])
sheet.add_row(['総計',        '',     '',     '=SUM(D2:D6)'])

# セルのマージ
sheet.merge_cells('A7:C7')

# 各行のフォントを変更
sheet.add_style('A1:D7', font_name: 'メイリオ')

# 背景色設定
sheet.add_style('A1:D1', bg_color: 'd0d0d0')
sheet.add_style('A7:A7', bg_color: 'b8cce4')

# # フォントカラー設定
sheet.add_style('D7:D7', fg_color: 'ae2f29')

# # 書体変更
sheet.add_style('D7:D7', b: true)

# アラインメント
sheet.add_style("A1:D1", alignment: { horizontal: :center })
sheet.add_style("A7:A7", alignment: { horizontal: :right })

# 罫線
sheet.add_style('A1:D7', border: { style: :thin, color: '000000' })

# 書き出し
package.serialize('output/example/caxlsx.xlsx')
