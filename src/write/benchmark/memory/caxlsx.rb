require_relative '../../../utils/memory'

SHEET_NUM = 100
ROW_NUM = 100
COL_NUM = 100

#==============================================================================

# メモリ使用量
print_memory_usage do
  require_relative '../../../utils/caxlsx'
  sheet, package = init_caxlsx(ROW_NUM, COL_NUM)
  sheet.add_style(
    range(0, 0, ROW_NUM - 1, COL_NUM - 1),
    font_name: 'メイリオ',
    bg_color: 'd0d0d0',
    fg_color: 'd0d0d0',
    b: true,
    alignment: { horizontal: :center },
    border: { style: :thin, color: '000000' }
  )

  package.serialize('output/benchmark/caxlsx.xlsx')
end
