# encoding: utf-8

Gem::Specification.new do |s|
  s.name = "docx_templater"
  s.version = File.read("lib/VERSION").strip

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.rubygems_version = "1.3.7"

  s.authors = ["Jonathan Andrew Wolter"]

  s.email = "jaw@jawspeak.com"

  s.date = "2011-12-10"
  s.description = "A Ruby library to template Microsoft Word .docx files."
  s.summary = "Generates new Word .docx files based on a template file."
  s.homepage = "https://github.com/jawspeak/ruby-docx-templater"

  s.rdoc_options = ["--charset=UTF-8"]
  s.extra_rdoc_files = ["README.rdoc"]

  s.require_paths = ["lib"]
  root_files = %w(docx_templater.gemspec LICENSE.txt Rakefile README.rdoc .gitignore Gemfile)
  s.files = Dir['{lib,script,spec}/**/*'] + root_files
  s.test_files = Dir['spec/**/*']

  s.add_dependency("nokogiri")
  # zipruby specifically because:
  #  - rubyzip does not support in-memory zip file modification (in you process sensitive info
  #  that can't hit the filesystem).
  #  - people report errors opening in word docx files when altered with rubyzip (search stackoverflow)

  # zipruby-compat, which is my fork of zipruby that changes the module Zip::* to ZipRuby::*,
  #  so it does not collide with rubyzip in same project.
  #  - see https://github.com/jawspeak/zipruby-compatibility-with-rubyzip-fork
  # THIS gem is not installed to rubyforge. Thus, you MUST on your own, in your Gemspec, add a dependency
  # to zipruby-compat. Example:
  #   gem "zipruby-compat", :git => "git@github.com:jawspeak/zipruby-compatibility-with-rubyzip-fork.git", :tag => "v0.3.7"
  s.add_dependency("zipruby-compat", "~> 0.3.7")

  s.add_development_dependency("rake")
  s.add_development_dependency("rspec", "~> 2.7.0")
end
