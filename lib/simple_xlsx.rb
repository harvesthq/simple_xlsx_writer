require 'tempfile'
require 'rubygems'

require 'zip/zip' #dep

unless String.method_defined? :to_xs
  require 'fast_xs' #dep
  class String
    alias_method :to_xs, :fast_xs
  end
end

$:.unshift(File.dirname(__FILE__))
require 'simple_xlsx/serializer'
require 'simple_xlsx/document'
require 'simple_xlsx/sheet'


