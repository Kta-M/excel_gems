require 'rubyXL'
require 'rubyXL/convenience_methods'

#==============================================================================
def init_rubyxl(row_num, col_num)
  workbook = RubyXL::Workbook.new
  sheet = workbook[0]
  add_cells_rubyxl(sheet, row_num, col_num)
  sheet
end

def add_cells_rubyxl(sheet, row_num, col_num)
  (0...row_num).each do |row|
    (0...col_num).each do |col|
      sheet.add_cell(row, col, 'hoge')
    end
  end
end
