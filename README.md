## Description

This is a simple, no-fuss generator for OpenXML (aka XLSX) files. No formatting or styles included -- just raw content with a few basic datatypes supported. Produced output is tested to be compatible with:

- Open Office 3.2 series (Linux, Mac, Windows)
- Neo Office 3.2 series (Mac)
- Microsoft Office 2007 (Windows)
- Microsoft Office 2010 (Windows)
- Microsoft Office 2008 for Mac (versions 12.2.5 or above)
- Microsoft Excel Viewer (Windows)
  
iWork's Numbers does not appear to support the inline string storage model used by this gem. Apple may release a fix for this eventually, but for now, I have avoided the more common shared string table method as it cannot be implemented in linear time.


## Sample
    
```ruby
require 'simple_xlsx'

serializer = SimpleXlsx::Serializer.new("test.xlsx") do |doc|
  doc.add_sheet("People") do |sheet|
    sheet.add_row(%w{DoB Name Occupation})
    sheet.add_row([Date.parse("July 31, 1912"), 
                  "Milton Friedman", 
                  "Economist / Statistician"])
  end
end
```

## License

See attached LICENSE for details.


## Credits

- Built by [Harvest](http://www.getharvest.com). Want to work on projects like this? [We’re hiring](http://www.getharvest.com/careers)!
- Concept and development by [Dee Zsombor](http://primalgrasp.com).
