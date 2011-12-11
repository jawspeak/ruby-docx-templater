require 'rspec'
require 'render_docx_template'
require 'template_processor_spec'

describe "integration test" do
  let(:data) { DocxTemplater::TestData::DATA }
  let(:input_file) { '../ExampleTemplate.docx' }
  let(:output_file) { 'spec/tmp/IntegrationTestOutput.docx' }
  before { File.delete(output_file) }

  context "should process in incoming docx" do
    it "generates a valid zip file (.docx)" do
        DocxCreator.new(input_file, data).generate_docx_file(output_file)

        archive = Zip::Archive.open(output_file)
        archive.close

        puts "************************************"
        puts "   >>> Only will work on mac <<<"
        puts "NOW attempting to open created file in Word."
        cmd = "open #{output_file}"
        puts "  will run '#{cmd}'"
        puts "************************************"

        system cmd
    end

    it "generates a file with the same contents as the input docx" do
      input_entries = Zip::Archive.open(input_file) { |z| z.map(&:name) }
      DocxCreator.new(input_file, data).generate_docx_file(output_file)
      output_entries = Zip::Archive.open(output_file) { |z| z.map(&:name) }

      input_entries.should == output_entries
    end
  end
end