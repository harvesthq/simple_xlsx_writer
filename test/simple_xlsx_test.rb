require File.dirname(__FILE__) + '/test_helper.rb'
require 'fileutils'

class SimpleXlsxTest < Test::Unit::TestCase

  def test_top_level
    FileUtils.rm_f "test.xlsx"
    o = SimpleXlsx::Serializer.new("test.xlsx") do |doc|
      doc.add_sheet "First" do |sheet|
        sheet.add_row ["Hello", "World", 3.14]
        sheet.add_row ["Another", "Row", Date.today]
      end
    end
  end

  if false
    def test_top_level_stream
      File.open "test_stream.xlsx", "wb" do |stream|
        o = SimpleXlsx::Serializer.new(stream) do |doc|
          doc.add_sheet "First" do |sheet|
            sheet.add_row ["Hello", "World", 3.14]
            sheet.add_row ["Another", "Row", Date.today]
          end
        end
      end
    end
  end

end
