module DocxTemplater
  module_function

  def log(str)
    # braindead logging
    puts str if ENV['DEBUG']
  end
end

require 'docx_templater/template_processor'
require 'docx_templater/docx_creator'
