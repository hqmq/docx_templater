require 'spec_helper'

describe Docx::PlaceholderObserver do
  describe "#next_node and #end_of_document" do
    subject{ Docx::PlaceholderObserver.new(data_provider) }
    let(:data_provider) do
      r = double("data_provider")
      allow(r).to receive(:has_key?).and_return(true)
      r
    end
    it "finds placeholders as it is given text nodes" do
      expect(data_provider).to receive(:[]).with(:title).and_return("The Thing")
      n1 = REXML::Text.new("dflkja sdf ||title|| slkjasdlkj")
      expect(n1).to receive(:value=).with('dflkja sdf The Thing slkjasdlkj')

      subject.next_node(n1)
      subject.end_of_document
    end

    it "finds placeholders among several nodes" do
      expect(data_provider).to receive(:[]).with(:title).and_return('The Thing')
      n1 = REXML::Text.new('booyah, (&&IJH))OJ |')
      expect(n1).to receive(:value=).with('booyah, (&&IJH))OJ The Thing')
      n2 = REXML::Text.new('|tit')
      expect(n2).to receive(:value=).with('')
      n3 = REXML::Text.new('le|| asdf093n38hfaj')
      expect(n3).to receive(:value=).with(' asdf093n38hfaj')

      subject.next_node(n1)
      subject.next_node(n2)
      subject.next_node(n3)
      subject.end_of_document
    end

    it "handles multiple placeholders in a single node correctly" do
      expect(data_provider).to receive(:[]).with(:title).and_return('Zombie Apocalypse')
      expect(data_provider).to receive(:[]).with(:subject).and_return('Movie')
      n1 = REXML::Text.new('||title|| is a ||subject||. Okay?')

      subject.next_node(n1)
      subject.end_of_document
      expect(n1.value).to eq 'Zombie Apocalypse is a Movie. Okay?'
    end

    it "handles nil replacements" do
      expect(data_provider).to receive(:[]).with(:title).and_return(nil)
      n1 = REXML::Text.new('||title||')
      subject.next_node(n1)
      subject.end_of_document
      expect(n1.value).to eq ''
    end
  end
end
