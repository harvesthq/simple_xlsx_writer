module SimpleXlsx

class Sheet
  attr_reader :name

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
      kind, ccontent = Sheet.format_field_and_type value
      row << "<c r=\"#{Sheet.column_index(col_ndx)}#{@row_ndx}\" t=#{kind}>#{ccontent}</c>"
    end
    row << "</row>"
    @row_ndx += 1
    @stream.write(row.join())
  end

  def self.format_field_and_type value
    if value.is_a?(String)
      [:inlineStr, "<is><t>#{value.to_xs}</t></is>"]
    elsif value.is_a?(Numeric)
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>"]
    elsif value.is_a?(Date)
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>"]
    elsif value.is_a?(DateTime)
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>"]
    elsif value.is_a?(TrueClass) || value.is_a?(FalseClass)
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>"]
    else
      [:inlineStr, "<is><t>#{value.to_s.to_xs}</t></is>"]
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
    result << abc[n]
    result.reverse.join
  end

end
end
