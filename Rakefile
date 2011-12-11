require 'rubygems'
require 'bundler'
Bundler.setup

require 'rake'
require 'rspec/core/rake_task'

RSpec::Core::RakeTask.new do |spec|
  spec.rspec_opts = ["--options", "spec/rspec.opts"]
end

task :default => :spec
