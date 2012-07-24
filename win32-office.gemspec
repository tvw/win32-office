# -*- encoding: utf-8 mode: ruby -*-
$:.push File.expand_path("../lib", __FILE__)
require "win32/office/version"

Gem::Specification.new do |s|
  s.name = "win32-office"
  s.version = Win32::Office::VERSION::STRING
  s.platform = Gem::Platform::RUBY
  s.authors = ["Thomas Volkmar Worm"]
  s.email = ["tvw@s4r.de"]
  s.homepage = "https://github.com/tvw/win32-office"
  s.summary = %q{A library for generating MS Office 2010 documents.}
  s.description = %q{A library for generating MS Office 2010 documents from document templates.}

  s.rubyforge_project = "win32-office"
  s.files = Dir["lib/**/*"] + ["Rakefile", "README.rdoc"]
  s.test_files = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]

  s.add_development_dependency 'rake', '>= 0.9.2'
end
