require 'objspace'

# メモリ使用量計測
def print_memory_usage
  memsize_before = ObjectSpace.memsize_of_all * 0.001 * 0.001
  rss_before = `ps -o rss= -p #{Process.pid}`.to_i * 0.001
  yield
  memsize_after = ObjectSpace.memsize_of_all * 0.001 * 0.001
  rss_after = `ps -o rss= -p #{Process.pid}`.to_i * 0.001
  puts "memsize_of_all: #{(memsize_after - memsize_before).round(2)} MB, rss: #{(rss_after - rss_before).round(2)} MB"
end
