require "test/unit"
require "rubygems"
require File.dirname(__FILE__) + '/../lib/simple_xlsx' unless defined?(SimpleXlsx)

require 'ruby-debug'
Debugger.settings[:autoeval] = true
Debugger.settings[:autolist] = 1
Debugger.start
