module DocxTemplater
  class TemplateProcessor
    attr_reader :data

    # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data)
      @data = data
    end

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

      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      log "begin_row_template: #{begin_row_template.to_s}"
      log "end_row_template: #{end_row_template.to_s}"
      raise "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_row_template && end_row_template

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

          begin_row_template.add_next_sibling(new_row)
        end
      end
      (row_templates + [begin_row_template, end_row_template]).map(&:unlink)
      xml.to_s
    end
  end
end