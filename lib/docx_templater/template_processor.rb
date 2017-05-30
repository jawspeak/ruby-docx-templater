require 'nokogiri'

module DocxTemplater
  class TemplateProcessor
    attr_reader :data, :escape_html, :skip_unmatched

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data, escape_html = true, skip_unmatched: false)
      @data = data
      @escape_html = escape_html
      @skip_unmatched = skip_unmatched
    end

    def render(document)
      document.force_encoding(Encoding::UTF_8) if document.respond_to?(:force_encoding)
      data.each do |key, value|
        case value
        when Array
          document = recursive_call_array(document, key, data[key])
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

    def enter_multiple_values xml, key, values
      xml = Nokogiri::XML(xml)

      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      unless begin_row_template && end_row_template
        return document if @skip_unmatched
        raise "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README."
      end

      row_templates = []
      row = begin_row_template.next_sibling
      while row != end_row_template
        row_templates.unshift(row)
        row = row.next_sibling
      end

      # for each data, reversed so they come out in the right order
      values.reverse_each do |data|
        rt = row_templates.map(&:dup)

        each_data = {}
        data.each do |k, v|
          if v.is_a?(Array)
            doc = Nokogiri::XML::Document.new
            root = doc.create_element 'pseudo_root', xml.root.namespaces
            root.inner_html = rt.reverse.map{|x| x.to_xml}.join
            q = recursive_call_array root.to_xml, k, v
            rt = xml.parse(q).reverse
          else
            each_data[k] = v
          end
        end


        # dup so we have new nodes to append
        rt.map(&:dup).each do |new_row|
          innards = new_row.inner_html
          matches = innards.scan(/\$EACH:([^\$]+)\$/)
          unless matches.empty?
            matches.map(&:first).each do |each_key|
              innards.gsub!("$EACH:#{each_key}$", each_data[each_key.downcase.to_s])
            end
          end
          # change all the internals of the new node, even if we did not template
          new_row.inner_html = innards
          #
          begin_row_template.add_next_sibling(new_row)
        end
      end
      (row_templates + [begin_row_template, end_row_template]).each(&:unlink)
      if xml.root.name == 'pseudo_root'
        xml.root.inner_html
      else
        xml.to_s
      end
    end

  end
end
