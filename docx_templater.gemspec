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
  root_files = %w(docx_templater.gemspec LICENSE.txt Rakefile README.rdoc .gitignore Gemfile)
  s.files = Dir['{lib,script,spec}/**/*'] + root_files
  s.test_files = Dir['spec/**/*']

  if RUBY_VERSION >= '1.9.2'
    s.add_runtime_dependency('nokogiri')
  else
    s.add_runtime_dependency('nokogiri', '~> 1.5.0')
  end

  # zipruby specifically because:
  #  - rubyzip does not support in-memory zip file modification (in you process sensitive info
  #  that can't hit the filesystem).
  #  - people report errors opening in word docx files when altered with rubyzip (search stackoverflow)
  s.add_runtime_dependency('zipruby')

  s.add_development_dependency('rake')
  s.add_development_dependency('rspec')
end
