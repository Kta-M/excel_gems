require 'benchmark'

require_relative '../../utils/spreadsheet'
require_relative '../../utils/rubyXL'
require_relative '../../utils/caxlsx'

SHEET_NUM = 100
ROW_NUM = 100
COL_NUM = 100

# Workbook作成
puts "\n## Create Workbook"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  x.report("spreadsheet") do
    Spreadsheet::Workbook.new
  end

  x.report("rubyXL") do
    RubyXL::Workbook.new
  end

  x.report("caxlsx") do
    package = Axlsx::Package.new
    package.workbook
  end
end

# Worksheet作成
puts "\n## Create Sheet"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  workbook = Spreadsheet::Workbook.new
  x.report("spreadsheet") do
    SHEET_NUM.times { |i| workbook.create_worksheet(name: i.to_s) }
  end

  workbook = RubyXL::Workbook.new
  x.report("rubyXL") do
    SHEET_NUM.times { |i| workbook.add_worksheet(i.to_s) }
  end

  package = Axlsx::Package.new
  workbook = package.workbook
  x.report("caxlsx") do
    SHEET_NUM.times { |i| workbook.add_worksheet(name: i.to_s) }
  end
end


# セルに書き込み
puts "\n## Add Cells"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  workbook = Spreadsheet::Workbook.new
  sheet = workbook.create_worksheet(name: 'sample')
  x.report("spreadsheet") do
    add_cells_spreadsheet(sheet, ROW_NUM, COL_NUM)
  end

  workbook = RubyXL::Workbook.new
  sheet = workbook[0]
  x.report("rubyXL") do
    add_cells_rubyxl(sheet, ROW_NUM, COL_NUM)
  end

  package = Axlsx::Package.new
  workbook = package.workbook
  sheet = workbook.add_worksheet(name: 'sample')
  x.report("caxlsx") do
    add_cells_caxlsx(sheet, ROW_NUM, COL_NUM)
  end
  package.serialize('output/benchmark/caxlsx.xlsx')
end

# セルのマージ
puts "\n## Merge Cells"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each_slice(2) do |row_begin, row_end|
      (0...COL_NUM).each do |col|
        sheet.merge_cells(row_begin, col, row_end, col)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each_slice(2) do |row_begin, row_end|
      (0...COL_NUM).each do |col|
        sheet.merge_cells(row_begin, col, row_end, col)
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    (0...ROW_NUM).each_slice(2) do |row_begin, row_end|
      (0...COL_NUM).each do |col|
        sheet.merge_cells(range(row_begin, col, row_end, col))
      end
    end
  end
end

# フォント変更
puts "\n## Change Font"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, font: Spreadsheet::Font.new('メイリオ'))
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_font_name('メイリオ')
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), font_name: 'メイリオ')
  end
end

# 背景色設定
puts "\n## Change Fill"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, pattern_fg_color: :silver, pattern: 1)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_fill('d0d0d0')
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), bg_color: 'd0d0d0')
  end
end

# フォントカラー設定
puts "\n## Change Font Color"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, color: :silver)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_font_color('d0d0d0')
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), fg_color: 'd0d0d0')
  end
end

# 書体設定
puts "\n## Change Font Color"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, weight: :bold)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_font_bold(true)
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), b: true)
  end
end

# アラインメント
puts "\n## Change Alignment"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, align: :center)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_horizontal_alignment('center')
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), alignment: { horizontal: :center })
  end
end

# 罫線
puts "\n## Change Border"
Benchmark.benchmark(Benchmark::CAPTION, 10, Benchmark::FORMAT) do |x|
  sheet = init_spreadsheet(ROW_NUM, COL_NUM)
  x.report("spreadsheet") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet.row(row).update_format(col, border: :thin)
      end
    end
  end

  sheet = init_rubyxl(ROW_NUM, COL_NUM)
  x.report("rubyXL") do
    (0...ROW_NUM).each do |row|
      (0...COL_NUM).each do |col|
        sheet[row][col].change_border(:top, 'thin')
        sheet[row][col].change_border(:left, 'thin')
        sheet[row][col].change_border(:bottom, 'thin')
        sheet[row][col].change_border(:right, 'thin')
      end
    end
  end

  sheet, _ = init_caxlsx(ROW_NUM, COL_NUM)
  x.report("caxlsx") do
    sheet.add_style(range(0, 0, ROW_NUM - 1, COL_NUM - 1), border: { style: :thin, color: '000000' })
  end
end
