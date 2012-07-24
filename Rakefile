#!/usr/bin/env rake
require "bundler/gem_tasks"
require 'rdoc/task'

task :default do
  sh %Q{bundle install}
  sh %Q{rake build}
  sh %Q{rake install}
  rm_rf "pkg"
end

Rake::RDocTask.new do |rd|
  rd.main = "README.rdoc"
  rd.rdoc_files.include("README.rdoc", "lib/**/*.rb")
  rd.rdoc_dir = "rdoc"
end
