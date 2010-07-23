module SimpleXlsx

class Serializer

  def initialize to
    @to = to
    Zip::ZipFile.open(to, Zip::ZipFile::CREATE) do |zip|
      @zip = zip
      add_doc_props
      add_relationship_part
      add_worksheets_directory
      @doc = Document.new(self)
      yield @doc
      add_content_types
      add_workbook_part
    end
  end

  def add_workbook_part
    @zip.get_output_stream "xl/workbook.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
ends
      @doc.sheets.each_with_index do |sheet, ndx|
        f.puts "<sheet name=\"#{sheet.name}\" sheetId=\"#{ndx + 1}\" r:id=\"rId#{ndx + 2}\"/>"
      end
      f.puts "</sheets></workbook>"
    end
  end

  def add_worksheets_directory
    @zip.mkdir "xl"
    @zip.mkdir "xl/worksheets"
  end

  def open_stream_for_sheet sheet_name
    @zip.get_output_stream "xl/worksheets/#{sheet_name}.xml" do |f|
      yield f
    end
  end

  def add_content_types
    @zip.get_output_stream "[Content_Types].xml" do |f|
      f.puts '<?xml version="1.0" encoding="UTF-8"?>'
      f.puts '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      f.puts <<-ends
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
ends
      if @doc.has_shared_strings?
        f.puts '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
      end
      @doc.sheets.each do |sheet|
        f.puts "<Override PartName=\"/xl/worksheets/#{sheet.name}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      end
      if @doc.has_styles?
        f.puts '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      end
      f.puts "</Types>"
    end
  end

  def add_relationship_part
    @zip.mkdir "_rels"
    @zip.get_output_stream "_rels/.rels" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
ends
    end
  end

  def add_doc_props
    @zip.mkdir "docProps"
    @zip.get_output_stream "docProps/core.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <dcterms:created xsi:type="dcterms:W3CDTF">2010-07-20T14:30:58.00Z</dcterms:created>
   <cp:revision>0</cp:revision>
</cp:coreProperties>
ends
    end
    @zip.get_output_stream "docProps/app.xml" do |f|
      f.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
</Properties>
ends
    end
  end

end

end


