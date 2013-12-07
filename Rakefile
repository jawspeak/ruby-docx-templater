require 'rspec/core/rake_task'
require 'rubocop/rake_task'

RSpec::Core::RakeTask.new(:spec)
Rubocop::RakeTask.new

task :default => [:spec, :rubocop]
