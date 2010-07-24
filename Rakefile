require 'rubygems'
require 'rake'
require 'rake/testtask'
require 'rake/gempackagetask'

task :default => [:test]

Rake::TestTask.new do |test|
  test.libs       << "test"
  test.test_files =  Dir['test/**/*_test.rb'].sort
  test.verbose    =  true
end

desc "generate tags for emacs"
task :tags do
  sh "ctags -Re lib/ "
end


# spec = Gem::Specification.new do |s|
#   s.name = "simple_xlsx_writer"
#   s.version = "0.0.1"
#   s.author = "Dee Zsombor"
#   s.email = "zsombor@getharvest.com"
#   s.homepage = "http://getharvest.com"
#   s.platform = Gem::Platform::RUBY
#   s.summary = "Just as the name says, simple writter for Office 2007+ Excel files"
#   s.files = FileList["{bin,lib}/**/*"].to_a
#   s.require_path = "lib"
#   s.autorequire = "name"
#   s.test_files = FileList["{test}/**/*test.rb"].to_a
#   s.has_rdoc = true
#   s.extra_rdoc_files = ["README"]
#   s.add_dependency("rubyzip", ">= 0.9.4")
# end

# Rake::GemPackageTask.new(spec) do |pkg|
#   pkg.need_tar = true
# end
