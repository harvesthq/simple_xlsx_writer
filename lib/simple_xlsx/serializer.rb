module SimpleXlsx

class Serializer

  def initialize file_path
    tempfile = Tempfile.new(File.basename(file_path))

    Zip::ZipOutputStream.open(tempfile.path) do |zip|
      @zip = zip
      add_doc_props
      add_relationship_part
      add_styles
      @doc = Document.new(self)
      yield @doc
      add_workbook_relationship_part
      add_content_types
      add_workbook_part
    end

    FileUtils.mkdir_p(File.dirname(file_path))
    FileUtils.cp(tempfile.path, file_path)
  end

  def add_workbook_part
    @zip.put_next_entry("xl/workbook.xml")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr date1904="0" />
<sheets>
ends
    @doc.sheets.each_with_index do |sheet, ndx|
      @zip.puts %Q{<sheet name="#{sheet.name}" sheetId="#{ndx + 1}" r:id="#{sheet.rid}"/>}
    end
    @zip.puts "</sheets></workbook>"
  end

  def open_stream_for_sheet ndx
    @zip.put_next_entry("xl/worksheets/sheet#{ndx + 1}.xml")
    @zip
  end

  def add_content_types
    @zip.put_next_entry("[Content_Types].xml")
    @zip.puts '<?xml version="1.0" encoding="UTF-8"?>'
    @zip.puts '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    @zip.puts <<-ends
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
ends
    if @doc.has_shared_strings?
      @zip.puts '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    end
    @doc.sheets.each_with_index do |sheet, ndx|
      @zip.puts %Q{<Override PartName="/xl/worksheets/sheet#{ndx+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>}
    end
    @zip.puts '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    @zip.puts "</Types>"
  end

  def add_workbook_relationship_part
    @zip.put_next_entry("xl/_rels/workbook.xml.rels")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
ends
    cnt = 0
    @zip.puts %Q{<Relationship Id="rId#{cnt += 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>}
    @doc.sheets.each_with_index do |sheet, ndx|
      sheet.rid = "rId#{cnt += 1}"
      @zip.puts %Q{<Relationship Id="#{sheet.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{ndx + 1}.xml"/>}
    end
    if @doc.has_shared_strings?
      @zip.puts '<Relationship Id="rId#{cnt += 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="xl/sharedStrings.xml"/>'
    end
    @zip.puts "</Relationships>"
  end

  def add_relationship_part
    @zip.put_next_entry("_rels/.rels")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
ends
    @zip.puts "</Relationships>"
  end

  def add_doc_props
    @zip.put_next_entry("docProps/core.xml")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
   <dcterms:created xsi:type="dcterms:W3CDTF">#{Time.now.utc.xmlschema}</dcterms:created>
   <cp:revision>0</cp:revision>
</cp:coreProperties>
ends
    @zip.put_next_entry("docProps/app.xml")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
</Properties>
ends
  end

  def add_styles
    @zip.put_next_entry("xl/styles.xml")
    @zip.puts <<-ends
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="7">
  <numFmt formatCode="GENERAL" numFmtId="164"/>
  <numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="170"/>
</numFmts>
<fonts count="5">
  <font><name val="Mangal"/><family val="2"/><sz val="10"/></font>
  <font><name val="Arial"/><family val="0"/><sz val="10"/></font>
  <font><name val="Arial"/><family val="0"/><sz val="10"/></font>
  <font><name val="Arial"/><family val="0"/><sz val="10"/></font>
  <font><name val="Arial"/><family val="2"/><sz val="10"/></font>
</fonts>
<fills count="2">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
  <border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="20">
  <xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">
    <alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>
    <protection hidden="false" locked="true"/>
  </xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"></xf>
  </cellStyleXfs>
<cellXfs count="7">
  <xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="164" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="22" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="15" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="1" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="2" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="49" xfId="0"></xf>
  <xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="4" numFmtId="170" xfId="0"></xf>
</cellXfs>
<cellStyles count="6"><cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>
  <cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>
  <cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>
  <cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>
  <cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>
  <cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>
</cellStyles>
</styleSheet>
ends
  end

end

end
