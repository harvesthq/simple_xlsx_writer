require File.dirname(__FILE__) + '/../test_helper.rb'

module SimpleXlsx

class SheetTest < Test::Unit::TestCase

  def test_column_index
    assert_equal 'A', Sheet.column_index(0)
    assert_equal 'B', Sheet.column_index(1)
    assert_equal 'C', Sheet.column_index(2)
    assert_equal 'D', Sheet.column_index(3)
    assert_equal 'Y', Sheet.column_index(24)
    assert_equal 'Z', Sheet.column_index(25)
  end

  def test_column_index_two_digits
    assert_equal 'AA', Sheet.column_index(0+26)
    assert_equal 'AB', Sheet.column_index(1+26)
    assert_equal 'AC', Sheet.column_index(2+26)
    assert_equal 'AD', Sheet.column_index(3+26)
    assert_equal 'AZ', Sheet.column_index(25+26)
    assert_equal 'BA', Sheet.column_index(25+26+1)
    assert_equal 'BB', Sheet.column_index(25+26+2)
    assert_equal 'BC', Sheet.column_index(25+26+3)
  end

  def test_header
    @sheet = Sheet.new(nil, 'something')
    puts @sheet.to_xml
    exp = <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>
</sheetData></worksheet>
ends
    assert_equal exp.strip, @sheet.to_xml
  end

end

end
