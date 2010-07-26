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


spec = Gem::Specification.new do |s|
  s.name = "simple_xlsx_writer"
  s.version = "0.5"
  s.author = "Dee Zsombor"
  s.email = "zsombor@primalgrasp.com"
  s.homepage = "http://simplxlsxwriter.rubyforge.org"
  s.rubyforge_project = "simple_xlsx_writer"
  s.platform = Gem::Platform::RUBY
  s.summary = "Just as the name says, simple writter for Office 2007+ Excel files"
  s.files = [FileList["{bin,lib}/**/*"].to_a, "LICENSE", "Rakefile"].flatten
  s.require_path = "lib"
  s.test_files = FileList["{test}/**/*test.rb"].to_a
  s.has_rdoc = true
  s.extra_rdoc_files = ["README"]
  s.add_dependency("rubyzip", ">= 0.9.4")
  s.add_dependency("fastxs", ">= 0.7.3")
end

Rake::GemPackageTask.new(spec) do |pkg|
  pkg.need_tar = true
end
