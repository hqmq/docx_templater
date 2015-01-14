require 'spec_helper'
require 'zip'
require 'tempfile'
describe DocxTemplater do
  let(:file_path){ File.expand_path("../../fixtures/TestFile.docx",__FILE__) }
  let(:spacing_file_path){ File.expand_path("../../fixtures/CorrectSpacing.docx",__FILE__) }
  
  describe "headers and footers" do
    let(:replacements){
      {:header => "Woohoo!", :date => "2012-01-01", :type => "Footsies" }
    }
    it "finds and replaces placeholders in the header" do
      str = get_header_string(file_path)
      expect(str).to include("||header||")
      expect(str).to_not include("Woohoo!")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_header_string(tf.path)
      expect(str).to_not include("||header||")
      expect(str).to include("Woohoo!")
    end

    it "finds and replaces placeholders in the footer" do
      str = get_footer_string(file_path)
      expect(str).to include("||date||")
      expect(str).to include("||type||")
      expect(str).to_not include("2012-01-01")
      expect(str).to_not include("Footsies")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_footer_string(tf.path)
      expect(str).to_not include("||date||")
      expect(str).to_not include("||type||")
      expect(str).to include("2012-01-01")
      expect(str).to include("Footsies")
    end
  end
  
  describe "body" do
    let(:replacements){
      {
        :title => "Working Title Please Ignore",
        :adjective => "FANTASTIC",
        :total_loan_amount_currency_words => "Three Hundred",
        :super_adjective => "BOOYAH",
        :non_string => 200.0,
        :side => "lefty",
        :by_side => "righty",
        :correct_spacing => "GIVE ME ROOM"
      }
    }
    it "finds and replaces placeholders in the body of the document" do
      str = get_body_string(file_path)
      expect(str).to include("||title||")
      expect(str).to include("||adjective||")
      expect(str).to_not include("Working Title Please Ignore")
      expect(str).to_not include("FANTASTIC")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      expect(str).to_not include("||title||")
      expect(str).to_not include("||adjective||")
      expect(str).to_not include("||total_loan_amount_currency_words||")
      expect(str).to include("Working Title Please Ignore")
      expect(str).to include("FANTASTIC")
      expect(str).to include("Three Hundred")
    end
    
    it "finds and replaces placeholders with formatting" do
      str = get_body_string(file_path)
      fragments = ['h ||','supe','r_adject','ive','| f']
      fragments.each do |fragment|
        expect(str).to include(fragment)
      end

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      fragments.each do |fragment|
        expect(str).to_not include(fragment)
      end
      expect(str).to include('BOOYAH')
    end

    it "finds and replaces with content that is not a string" do
      str = get_body_string(file_path)
      expect(str).to include("||non_string||")
      expect(str).to_not include("200.0")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      expect(str).to_not include("||non_string||")
      expect(str).to include("200.0")
    end

    it "If no data provider key matches, it should leave the placeholder" do
      str = get_body_string(file_path)
      expect(str).to include("||stay_on_the_page")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      expect(str).to include("||stay_on_the_page")
    end

    it "should handle side by side placeholders" do
      str = get_body_string(file_path)
      expect(str).to include("side")
      expect(str).to include("by_side")

      buffer = ::DocxTemplater.new.replace_file_with_content( file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      expect(str).to_not include("side")
      expect(str).to_not include("by_side")
      expect(str).to include("lefty")
      expect(str).to include("righty")
    end

    it "should correctly preserve spacing before and after placeholders" do
      replacements[:placeholders] = "space"
      replacements[:with_spaces] = "balls"
      str = get_body_string(spacing_file_path)
      expect(str).to include("Test ||placeholders|| ||with_spaces|")
      expect(str).to include("<w:t>|      body time</w:t>")

      buffer = ::DocxTemplater.new.replace_file_with_content( spacing_file_path, replacements )
      tf = Tempfile.new(["spec","docx"])
      tf.write buffer.string
      tf.close

      str = get_body_string(tf.path)
      expect(str).to_not include("Test ||placeholders|| ||with_spaces|")
      expect(str).to_not include("|      body time")
      expect(str).to include("Test space balls")
      expect(str).to include("<w:t xml:space='preserve'>      body time</w:t>")
    end

  end
end
