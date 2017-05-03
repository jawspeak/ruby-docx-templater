# encoding: utf-8

Gem::Specification.new do |s|
  s.name = 'docx_templater'
  s.version = File.read('lib/VERSION').strip

  s.authors = ['Jonathan Andrew Wolter']

  s.email = 'jaw@jawspeak.com'

  s.date = '2011-12-10'
  s.description = 'A Ruby library to template Microsoft Word .docx files.'
  s.summary = 'Generates new Word .docx files based on a template file.'
  s.homepage = 'https://github.com/jawspeak/ruby-docx-templater'

  s.rdoc_options = ['--charset=UTF-8']
  s.extra_rdoc_files = ['README.rdoc']

  s.require_paths = ['lib']
  root_files = %w[docx_templater.gemspec LICENSE.txt Rakefile README.rdoc .gitignore Gemfile]
  s.files = Dir['{lib,script,spec}/**/*'] + root_files
  s.test_files = Dir['spec/**/*']

  s.add_runtime_dependency('nokogiri')
  s.add_runtime_dependency('rubyzip', '>= 1.1.1')

  s.add_development_dependency('rake')
  s.add_development_dependency('rubocop')
  s.add_development_dependency('rspec')
end
