require 'spec_helper'
describe DocxTemplater do
  subject{ DocxTemplater.new }
  describe "#generate_tags_for" do
    let(:m1){
      m1 = double("model")
      allow(m1).to receive(:attributes).and_return({mikey: "snuggy", ruby: "awesome"})
      m1
    }
    let(:m2){
      m2 = double("model")
      allow(m2).to receive(:attributes).and_return({paul: "smith", ruby: "amazing"})
      m2
    }

    it "combines multiple hashes" do
      h1 = {foo: "foo", bar: "bar"}
      h2 = {baz: "baz"}

      combined = subject.generate_tags_for(h1, h2)
      expect(combined[:foo]).to eq "foo"
      expect(combined[:bar]).to eq "bar"
      expect(combined[:baz]).to eq "baz"
    end

    it "can combine multiple #attributes models" do
      #note this should be working, but we have an invalid dependency on ActiveSupport
      # to make String#underscore available
      expect {
        combined = subject.generate_tags_for(m1, m2)
        expect(combined[:mikey]).to eq "snuggy"
        expect(combined[:paul]).to eq "smith"
      }.to raise_error(NoMethodError)
    end

    it "can namespace #attributes models" do
      combined = subject.generate_tags_for(
        {:prefix => "m1", :data => m1},
        {:prefix => "m2", :data => m2})
      expect(combined[:m1_mikey]).to eq "snuggy"
      expect(combined[:m2_paul]).to eq "smith"
      expect(combined[:m1_ruby]).to eq "awesome"
      expect(combined[:m2_ruby]).to eq "amazing"
    end

    it "prefers to call #template_attributes for easy customization" do
      allow(m1).to receive(:template_attributes).and_return({mikey: "he likes it", ruby: "fast"})
      combined = subject.generate_tags_for(
        {prefix: "m1", :data => m1},
        {prefix: "m2", :data => m2})
      expect(combined[:m1_mikey]).to eq "he likes it"
      expect(combined[:m1_ruby]).to eq "fast"
    end
  end
end
