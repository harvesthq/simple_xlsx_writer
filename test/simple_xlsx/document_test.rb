require File.join('.', File.dirname(__FILE__), '..', 'test_helper')

module SimpleXlsx

class DocumentTest < Test::Unit::TestCase

  def open_stream_for_sheet sheets_size
    assert_equal sheets_size, @doc.sheets.size
    self
  end

  def write arg
  end

  def test_add_sheet
    @doc = Document.new self
    assert_equal [], @doc.sheets
    @doc.add_sheet "new sheet"
    assert_equal 1, @doc.sheets.size
    assert_equal 'new sheet', @doc.sheets.first.name
  end

end
end
