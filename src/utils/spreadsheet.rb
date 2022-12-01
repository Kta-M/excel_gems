require 'spreadsheet'

#==============================================================================
def init_spreadsheet(row_num, col_num)
  workbook = Spreadsheet::Workbook.new
  sheet = workbook.create_worksheet(name: 'Sheet1')
  add_cells_spreadsheet(sheet, row_num, col_num)
  sheet
end

def add_cells_spreadsheet(sheet, row_num, col_num)
  (0...row_num).each do |row|
    (0...col_num).each do |col|
      sheet[row, col] = 'hoge'
    end
  end
end
