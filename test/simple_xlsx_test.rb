require File.join('.', File.dirname(__FILE__), 'test_helper')
require 'fileutils'

class SimpleXlsxTest < Test::Unit::TestCase

  def setup
    FileUtils.rm_f "test.xlsx"
  end

  def teardown
    FileUtils.rm_f "test.xlsx"
  end

  def test_top_level
    o = SimpleXlsx::Serializer.new("test.xlsx") do |doc|
      doc.add_sheet "First" do |sheet|
        sheet.add_row ["Hello", "World", 3.14, 7]
        sheet.add_row ["Another", "Row", Date.today, Time.parse('2010-Jul-24 12:00 UTC')]
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
