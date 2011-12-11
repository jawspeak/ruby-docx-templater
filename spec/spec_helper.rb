require 'rubygems'
require 'bundler'
Bundler.setup

require 'rspec'
require 'nokogiri'
require 'fileutils'

require 'docx_templater'

SPEC_BASE_PATH = Pathname.new(File.expand_path(File.dirname(__FILE__)))
