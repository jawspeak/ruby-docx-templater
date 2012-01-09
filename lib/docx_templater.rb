require 'rubygems'
require 'nokogiri'
require 'zipruby'

module DocxTemplater
  def log(str)
    # braindead logging
    # puts str
  end
  extend self
end

require 'docx_templater/template_processor'
require 'docx_templater/docx_creator'
