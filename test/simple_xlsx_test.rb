require File.dirname(__FILE__) + '/test_helper.rb'
require 'fileutils'

class SimpleXlsxTest < Test::Unit::TestCase

  def test_top_level
    FileUtils.rm_f "test.xmlx"
    o = SimpleXlsx::Serializer.new("test.xmlx") do |doc|
      doc.add_sheet "First" do |sheet|
        sheet.add_row ["Hello", "World", 3.14]
        sheet.add_row ["Another", "Row", Date.today]
      end
    end
  end

end
