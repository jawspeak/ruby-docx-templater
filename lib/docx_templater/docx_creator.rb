require 'zip'

module DocxTemplater
  class DocxCreator
    attr_reader :template_path, :template_processor

    def initialize(template_path, data, escape_html = true)
      @template_path = template_path
      @template_processor = TemplateProcessor.new(data, escape_html)
    end

    def generate_docx_file(file_name = "output_#{Time.now.strftime('%Y-%m-%d_%H%M')}.docx")
      File.open(file_name, 'w') do |f|
        f.write(generate_docx_bytes.string)
      end
    end

    def generate_docx_bytes
      Zip::OutputStream.write_buffer(StringIO.new) do |out|
        Zip::File.open(template_path).each do |entry|
          entry_name = entry.name
          out.put_next_entry(entry_name)
          out.write(copy_or_template(entry_name, entry.get_input_stream.read))
        end
      end
    end

    private

    def copy_or_template(entry_name, entry_bytes)
      # Inside the word document archive is one file with contents of the actual document. Modify it.
      if entry_name == 'word/document.xml'
        template_processor.render(entry_bytes)
      else
        entry_bytes
      end
    end
  end
end
