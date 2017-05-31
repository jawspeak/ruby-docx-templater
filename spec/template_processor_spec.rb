# encoding: UTF-8

require 'spec_helper'
require 'nokogiri'

module DocxTemplater
  module TestData
    DATA = {
      teacher: 'Priya Vora',
      building: 'Building #14',
      classroom: 'Rm 202'.to_sym,
      district: 'Washington County Public Schools',
      senority: 12.25,
      roster: [
        { name: 'Sally', age: 12, attendence: '100%' },
        { name: :Xiao, age: 10, attendence: '94%' },
        { name: 'Bryan', age: 13, attendence: '100%' },
        { name: 'Larry', age: 11, attendence: '90%' },
        { name: 'Kumar', age: 12, attendence: '76%' },
        { name: 'Amber', age: 11, attendence: '100%' },
        { name: 'Isaiah', age: 12, attendence: '89%' },
        { name: 'Omar', age: 12, attendence: '99%' },
        { name: 'Xi', age: 11, attendence: '20%' },
        { name: 'Noushin', age: 12, attendence: '100%' }
      ],
      event_reports: [
        { name: 'Science Museum Field Trip', notes: 'PTA sponsored event. Spoke to Astronaut with HAM radio.' },
        { name: 'Wilderness Center Retreat', notes: '2 days hiking for charity:water fundraiser, $10,200 raised.' }
      ],
      true_cond: true,
      false_cond: false,
      created_at: '11-12-03 02:01'
    }.freeze
  end
end

describe DocxTemplater::TemplateProcessor do
  let(:data) { Marshal.load(Marshal.dump(DocxTemplater::TestData::DATA)) } # deep copy
  let(:base_path) { SPEC_BASE_PATH.join('example_input') }
  let(:xml) { File.read("#{base_path}/word/document.xml") }
  let(:parser) { DocxTemplater::TemplateProcessor.new(data) }

  context 'valid xml' do
    it 'should render and still be valid XML' do
      expect(Nokogiri::XML.parse(xml)).to be_xml
      out = parser.render(xml)
      expect(Nokogiri::XML.parse(out)).to be_xml
    end

    it 'should accept non-ascii characters' do
      data[:teacher] = '老师'
      out = parser.render(xml)
      expect(out).to include('老师')
      expect(Nokogiri::XML.parse(out)).to be_xml
    end

    it 'should escape as necessary invalid xml characters, if told to' do
      data[:building] = '23rd & A #1 floor'
      data[:classroom] = '--> 201 <!--'
      data[:roster][0][:name] = '<#Ai & Bo>'
      out = parser.render(xml)

      expect(Nokogiri::XML.parse(out)).to be_xml
      expect(out).to include('23rd &amp; A #1 floor')
      expect(out).to include('--&gt; 201 &lt;!--')
      expect(out).to include('&lt;#Ai &amp; Bo&gt;')
    end

    context 'not escape xml' do
      let(:parser) { DocxTemplater::TemplateProcessor.new(data, false) }
      it 'does not escape the xml attributes' do
        data[:building] = '23rd <p>&amp;</p> #1 floor'
        out = parser.render(xml)
        expect(Nokogiri::XML.parse(out)).to be_xml
        expect(out).to include('23rd <p>&amp;</p> #1 floor')
      end
    end
    context 'when unmatched array' do
      let(:unmatched_data) do
        data.merge(unmatched_array: ['Some data'])
      end

      it 'raised' do
        parser = DocxTemplater::TemplateProcessor.new(unmatched_data)
        expect { parser.render(xml) }.to raise_error(RuntimeError)
      end

      it 'skipped' do
        parser = DocxTemplater::TemplateProcessor.new(unmatched_data, skip_unmatched: true)
        out = parser.render(xml)
        expect(Nokogiri::XML.parse(out)).to be_xml
      end
    end
  end

  context 'unmatched begin and end row templates' do
    it 'should not raise' do
      xml = <<EOF
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#BEGIN_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#END_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#BEGIN_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#END_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
    </w:tbl>
  </w:body>
</xml>
EOF
      expect { parser.render(xml) }.to_not raise_error
    end

    it 'should raise an exception' do
      xml = <<EOF
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#BEGIN_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#END_ROW:#{:roster.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
      <w:tr><w:tc>
          <w:p>
            <w:r><w:t>#BEGIN_ROW:#{:event_reports.to_s.upcase}#</w:t></w:r>
          </w:p>
      </w:tc></w:tr>
    </w:tbl>
  </w:body>
</xml>
EOF
      expect { parser.render(xml) }.to raise_error(/#END_ROW:EVENT_REPORTS# nil: true/)
    end
  end

  it 'should enter no text for a nil value' do
    xml = <<EOF
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
  <w:p>Before.$KEY$After</w:p>
</w:body>
</xml>
EOF
    actual = DocxTemplater::TemplateProcessor.new(key: nil).render(xml)
    expected_xml = <<EOF
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
  <w:p>Before.After</w:p>
</w:body>
</xml>
EOF
    expect(actual).to eq(expected_xml)
  end

  it 'should replace all simple keys with values' do
    non_array_keys = data.reject { |_, v| [Array, TrueClass, FalseClass].include?(v.class) }
    non_array_keys.keys.each do |key|
      expect(xml).to include("$#{key.to_s.upcase}$")
      expect(xml).not_to include(data[key].to_s)
    end
    out = parser.render(xml)

    non_array_keys.each do |key|
      expect(out).not_to include("$#{key}$")
      expect(out).to include(data[key].to_s)
    end
  end

  it 'should replace all array keys with values' do
    expect(xml).to include('#BEGIN_ROW:')
    expect(xml).to include('#END_ROW:')
    expect(xml).to include('$EACH:')

    out = parser.render(xml)

    expect(out).not_to include('#BEGIN_ROW:')
    expect(out).not_to include('#END_ROW:')
    expect(out).not_to include('$EACH:')

    %i[roster event_reports].each do |key|
      data[key].each do |row|
        row.values.map(&:to_s).each do |row_value|
          expect(out).to include(row_value)
        end
      end
    end
  end

  it 'should replace all condition keys' do
    expect(xml).to include('#IF:')
    expect(xml).to include('#ENDIF:')

    out = parser.render(xml)

    expect(out).not_to include('#IF:')
    expect(out).not_to include('#ENDIF:')
    expect(out).not_to include('dont show this')
  end

  it 'shold render students names in the same order as the data' do
    out = parser.render(xml)
    expect(out).to include('Sally')
    expect(out).to include('Kumar')
    expect(out.index('Kumar')).to be > out.index('Sally')
  end

  it 'shold render event reports names in the same order as the data' do
    out = parser.render(xml)
    expect(out).to include('Science Museum Field Trip')
    expect(out).to include('Wilderness Center Retreat')
    expect(out.index('Wilderness Center Retreat')).to be > out.index('Science Museum Field Trip')
  end

  it 'should render 2-line event reports in same order as docx' do
    event_reports_starting_at = xml.index('#BEGIN_ROW:EVENT_REPORTS#')
    expect(event_reports_starting_at).to be >= 0
    expect(xml.index('$EACH:NAME$', event_reports_starting_at)).to be > event_reports_starting_at
    expect(xml.index('$EACH:NOTES$', event_reports_starting_at)).to be > event_reports_starting_at
    expect(xml.index('$EACH:NOTES$', event_reports_starting_at)).to be > xml.index('$EACH:NAME$', event_reports_starting_at)

    out = parser.render(xml)
    expect(out.index('PTA sponsored event. Spoke to Astronaut with HAM radio.')).to be > out.index('Science Museum Field Trip')
  end

  it 'should render sums of input data' do
    expect(xml).to include('#SUM')
    out = parser.render(xml)
    expect(out).not_to include('#SUM')
    expect(out).to include("#{data[:roster].count} Students")
    expect(out).to include("#{data[:event_reports].count} Events")
  end
end
