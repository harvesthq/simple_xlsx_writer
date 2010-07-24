require 'bigdecimal'

module SimpleXlsx

class Sheet
  attr_reader :name
  attr_accessor :rid

  def initialize document, name, stream, &block
    @document = document
    @stream =  stream
    @name = name
    @row_ndx = 1
    @stream.write <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
ends
    if block_given?
      yield self
    end
    @stream.write "</sheetData></worksheet>"
  end

  def add_row arry
    row = ["<row r=\"#{@row_ndx}\">"]
    arry.each_with_index do |value, col_ndx|
      kind, ccontent, cstyle = Sheet.format_field_and_type_and_style value
      row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" t=\"#{kind.to_s}\" s=\"#{cstyle}\">#{ccontent}</c>"
    end
    row << "</row>"
    @row_ndx += 1
    @stream.write(row.join())
  end

  def self.format_field_and_type_and_style value
    if value.is_a?(String)
      [:inlineStr, "<is><t>#{value.to_xs}</t></is>", 5]
    elsif value.is_a?(BigDecimal)
      [:n, "<v>#{value.to_s('f')}</v>", 3]
    elsif value.is_a?(Numeric)
      [:n, "<v>#{value.to_s}</v>", 3]
    elsif value.is_a?(Date)
      [:inlineStr, "<is><t>#{value.strftime('%Y-%b-%d')}</t></is>", 1]
    elsif value.is_a?(DateTime)
      [:inlineStr, "<is><t>#{value.strftime('%Y-%b-%d %H:%M:%S')}</t></is>", 2]
    elsif value.is_a?(TrueClass) || value.is_a?(FalseClass)
      [:b, "<v>#{value ? '1' : '0'}</v>", 6]
    else
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>", 5]
    end
  end

  def self.abc
    @@abc ||= ('A'..'Z').to_a
  end

  def self.column_index n
    result = []
    while n >= 26 do
      result << abc[n % 26]
      n /= 26
    end
    result << abc[result.empty? ? n : n - 1]
    result.reverse.join
  end

end
end
