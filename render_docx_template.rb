require 'rubygems'
require 'nokogiri'
require 'zipruby'

def log(str)
  # puts str
end

module DocxTemplater
  class TemplateProcesser
    attr_reader :data

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data)
      @data = data
    end

    # naive and innefficient templating.
    def render(document)
      data.each do |key, value|
        if value.class == Array
          document = enter_multiple_values(document, key)
          document.gsub!("#SUM:#{key.to_s.upcase}#", value.count.to_s)
        else
          document.gsub!("$#{key.to_s.upcase}$", safe(value))
        end
      end
      document
    end

    private

    def safe(text)
      text.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
    end

    def enter_multiple_values(document, key)
      log "enter_multiple_values for: #{key}"
      # TODO ideally we would not re-parse xml doc every time
      xml = Nokogiri::XML(document)

      # Often these tags in Word are broken up with various xml entries. Probably need to manually fix word/document.xml before saving the template.
      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      log "begin_row_template: #{begin_row_template.to_s}"
      log "end_row_template: #{end_row_template.to_s}"
      raise "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word messed with format, or xml became invalid. See README." unless begin_row_template && end_row_template

      row_templates = []
      row = begin_row_template.next_sibling
      while (row != end_row_template)
        row_templates.unshift(row)
        row = row.next_sibling
      end
      log "row_templates: (#{row_templates.count}) #{row_templates.map(&:to_s).inspect}"

      # for each data, reversed so they come out in the right order
      data[key].reverse.each do |each_data|
        log "each_data: #{each_data.inspect}"

        # dup so we have new nodes to append
        row_templates.map(&:dup).each do |new_row|
          log "   new_row: #{new_row}"
          innards = new_row.inner_html
          if !(matches = innards.scan(/\$EACH:([^\$]+)\$/)).empty?
            log "   matches: #{matches.inspect}"
            matches.map(&:first).each do |each_key|
              log "      each_key: #{each_key}"
              innards.gsub!("$EACH:#{each_key}$", safe(each_data[each_key.downcase.to_sym]))
            end
          end
          # change all the internals of the new node, even if we did not template
          new_row.inner_html = innards
          #log "new_row new innards: #{new_row.inner_html}"

          # add this row after the template's start
          begin_row_template.add_next_sibling(new_row)
        end
      end
      # delete unwanted template rows from document
      (row_templates + [begin_row_template, end_row_template]).map(&:unlink)
      xml.to_s
    end
  end

  # Creates a new word document from an existing docx file. (You may need to modify that docx since word
  # may munge your templating markup with in-between XML nodes.)
  class DocxCreator
    attr_reader :template_path, :data, :template_parser

    def initialize(template_path, data)
      @template_path = template_path
      @template_parser = TemplateProcesser.new(data)
    end

    def generate_docx_file(file_name = "output_#{Time.now.strftime("%Y-%m-%d_%H%M")}.docx")
      buffer = generate_docx_bytes
      File.open(file_name, 'w') { |f| f.write(buffer) }
    end

    def generate_docx_bytes
      buffer = ''

      # Open the existing template file (no temp files created, just read it)
      Zip::Archive.open(template_path) do |template|
        n_entries = template.num_files

        # Then create a new file with the output kept in-memory.
        Zip::Archive.open_buffer(buffer, Zip::CREATE) do |archive|
          n_entries.times do |i|
            entry_name = template.get_name(i)
            template.fopen(entry_name) do |f|
              archive.add_buffer(entry_name, copy_or_template(entry_name, f))
            end
          end
        end
      end
      buffer
    end

    private

    def copy_or_template(entry_name, f)
      # Inside the word document archive is one file with contents of the actual document. Modify it.
      return template_parser.render(f.read) if entry_name == 'word/document.xml'
      f.read
    end
  end
end
