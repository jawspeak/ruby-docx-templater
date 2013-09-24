module DocxTemplater
  extend self
  def log(str)
    # braindead logging
    # puts str
  end
end

require 'docx_templater/template_processor'
require 'docx_templater/docx_creator'
