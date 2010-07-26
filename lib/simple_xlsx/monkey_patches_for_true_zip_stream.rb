require 'zip/zip' #dep

__END__

module Zip
  class ZipOutputStream
    def initialize(fileName)
      super()
      if fileName.is_a?(String) && !fileName.empty?
        @fileName = fileName
        @outputStream = File.new(@fileName, "wb")
      else
        @outputStream = fileName
        @fileName = ''
      end
      @entrySet = ZipEntrySet.new
      @compressor = NullCompressor.instance
      @closed = false
      @currentEntry = nil
      @comment = nil
    end
  end

  class ZipFile < ZipCentralDirectory
    def initialize(stream, create = nil)
      super()
      @name = stream.is_a?(String) ? stream : ''
      @comment = ""
      if stream.is_a?(String) && File.exists?(stream)
        File.open(name, "rb") { |f| read_from_stream(f) }
      elsif (create)
        @entrySet = ZipEntrySet.new
      elsif !stream.is_a?(String) && !create && !stream.respond_to(:path)
        # do nothing here
      elsif !stream.is_a?(String) && !create
        File.open(stream.path, "rb") { |f| read_from_stream(f) }
      else
        raise ZipError, "File #{stream} not found"
      end
      @create = create
      @storedEntries = @entrySet.dup

      @restore_ownership = false
      @restore_permissions = false
      @restore_times = true
    end

    def on_success_replace arg
      if arg.is_a?(String) && !arg.empty?
        tmpfile = get_tempfile
        tmpFilename = tmpfile.path
        tmpfile.close
        if yield tmpFilename
          File.rename(tmpFilename, name)
        end
      else
        yield arg
      end
    end
  end
end
