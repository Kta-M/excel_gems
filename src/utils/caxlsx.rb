require 'caxlsx'

#==============================================================================
def init_caxlsx(row_num, col_num)
  package = Axlsx::Package.new
  workbook = package.workbook
  sheet = workbook.add_worksheet(name: 'Sheet1')
  add_cells_caxlsx(sheet, row_num, col_num)
  [sheet, package]
end

def add_cells_caxlsx(sheet, row_num, col_num)
  (0...row_num).each do |row|
    sheet.add_row(['hoge'] * col_num)
  end
end

# セルの範囲をExcelのレンジフォーマットで返す
def range(row_begin, col_begin, row_end, col_end)
  "#{index_to_col_name(col_begin)}#{row_begin+1}:#{index_to_col_name(col_end)}#{row_end+1}"
end

# 列のindex値を列名で返す
def index_to_col_name(index)
  offset = 'A'.ord
  result = []

  val = index + 1
  while val > 0
    val -= 1
    result.unshift((val % 26 + offset).chr)
    val /= 26
  end
  result.join
end

