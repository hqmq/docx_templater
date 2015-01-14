require 'spec_helper'
describe Docx::DocumentReplacer do
  subject{ Docx::DocumentReplacer.new(xml_str, data_provider)}
  let(:xml_str){ File.read( File.expand_path('spec/fixtures/header2.xml') )}
  let(:data_provider) do
  	r = double("data_provider", :[] => "Mikey Header")
  	allow(r).to receive(:has_key?).and_return(true)
    r
  end
  it "walks an xml string and replaces values" do
    expect(xml_str).to include('||header||')

    expect(subject.replaced).not_to include('||header||')
    expect(subject.replaced).to include('Mikey Header')
  end
end
