require 'nokogiri'

module DocxTemplater
  class TemplateProcessor
    attr_reader :data, :escape_html

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data, escape_html = true)
      @data = data
      @escape_html = escape_html
    end

    def render(document)
      document.force_encoding(Encoding::UTF_8) if document.respond_to?(:force_encoding)
      data.each do |key, value|
        case value
        when Array
          document = enter_multiple_values(document, key)
          document.gsub!("#SUM:#{key.to_s.upcase}#", value.count.to_s)
        when TrueClass, FalseClass
          if value
            document.gsub!(/\#(END)?IF:#{key.to_s.upcase}\#/, '')
          else
            document.gsub!(/\#IF:#{key.to_s.upcase}\#.*\#ENDIF:#{key.to_s.upcase}\#/m, '')
          end
        else
          document.gsub!("$#{key.to_s.upcase}$", safe(value))
        end
      end
      document
    end

    private

    def safe(text)
      if escape_html
        text.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
      else
        text.to_s
      end
    end

    def enter_multiple_values(document, key)
      DocxTemplater.log("enter_multiple_values for: #{key}")
      # TODO: ideally we would not re-parse xml doc every time
      xml = Nokogiri::XML(document)

      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      DocxTemplater.log("begin_row_template: #{begin_row_template}")
      DocxTemplater.log("end_row_template: #{end_row_template}")
      fail "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_row_template && end_row_template

      row_templates = []
      row = begin_row_template.next_sibling
      while row != end_row_template
        row_templates.unshift(row)
        row = row.next_sibling
      end
      DocxTemplater.log("row_templates: (#{row_templates.count}) #{row_templates.map(&:to_s).inspect}")

      # for each data, reversed so they come out in the right order
      data[key].reverse_each do |each_data|
        DocxTemplater.log("each_data: #{each_data.inspect}")

        # dup so we have new nodes to append
        row_templates.map(&:dup).each do |new_row|
          DocxTemplater.log("   new_row: #{new_row}")
          innards = new_row.inner_html
          matches = innards.scan(/\$EACH:([^\$]+)\$/)
          unless matches.empty?
            DocxTemplater.log("   matches: #{matches.inspect}")
            matches.map(&:first).each do |each_key|
              DocxTemplater.log("      each_key: #{each_key}")
              innards.gsub!("$EACH:#{each_key}$", safe(each_data[each_key.downcase.to_sym]))
            end
          end
          # change all the internals of the new node, even if we did not template
          new_row.inner_html = innards
          # DocxTemplater::log("new_row new innards: #{new_row.inner_html}")

          begin_row_template.add_next_sibling(new_row)
        end
      end
      (row_templates + [begin_row_template, end_row_template]).each(&:unlink)
      xml.to_s
    end
  end
end
