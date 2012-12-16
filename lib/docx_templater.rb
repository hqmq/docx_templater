require 'zip/zipfilesystem'
require 'htmlentities'

# Use .docx as reusable templates
# 
# Example:
# buffer = DocxTemplater.replace_file_with_content('path/to/mydocument.docx',
#    {
#      :client_email1 => 'test@example.com',
#      :client_phone1 => '555-555-5555',
#    })
# # In Rails you can send a word document via send_data
# send_data buffer.string, :filename => 'REPC.docx'
# # Or save the output to a word file
# File.open("path/to/mydocument.docx", "wb") {|f| f.write(buffer.string) }
class DocxTemplater
  def initialize(opts = {})
    @options = opts
  end

  def replace_file_with_content(file_path, data_provider)
    # Rubyzip doesn't save it right unless saved like this: https://gist.github.com/e7d2855435654e1ebc52
    zf = Zip::ZipFile.new(file_path) # Put original file name here

    buffer = Zip::ZipOutputStream.write_buffer do |out|
      zf.entries.each do |e|
        process_entry(e, out, data_provider)
      end
    end
    # You can save this buffer or send it with rails via send_data
    return buffer
  end

  def generate_tags_for(*args)
    attributes = {}
    args.flatten!
    # Prefixes the model name or custom prefix. Makes it so we don't having naming clashes when used with records from multiple m
    args.each do |arg|
      if arg.is_a?(Hash) && arg.has_key?(:data) && arg.has_key?(:prefix)
        template_attributes = (arg[:data].respond_to?(:template_attributes) && :template_attributes) || :attributes
        arg[:data].send(template_attributes).each_key do |key|
          attributes["#{arg[:prefix]}_#{key.to_s}".to_sym] = arg[:data].send(template_attributes)[key]
        end
      elsif arg.is_a?(Hash)
        attributes.merge!(arg)
      else
        template_attributes = (arg.respond_to?(:template_attributes) && :template_attributes) || :attributes
        arg.send(template_attributes).each_key do |key|
          attributes["#{arg.class.name.underscore}_#{key.to_s}".to_sym] = arg.send(template_attributes)[key]
        end
      end
    end
    attributes
  end

  private
  
  def all_tags_regex
    /\|\|\<*.+?\>*\|\|/
  end
  
  def malformed_tag_regex
    /(?<=>)\w{3,}(?=<)/
  end
  
  def well_formed_tag_regex
    /(?<=\|\|)\w{3,}(?=\|\|)/
  end
  
  def just_label_regex
    /(?<=>)(\w{3,})/
  end
  
  def entry_requires_replacement?(entry)
    entry.ftype != :directory && entry.name =~ /document|header|footer/
  end
  
  def get_entry_content(entry, data_provider)
    if entry_requires_replacement?(entry)
      replace_entry_content(entry.get_input_stream.read, data_provider)
    else
      entry.get_input_stream.read
    end
  end
  
  def process_entry(entry, output, data_provider)
    output.put_next_entry(entry.name)
    output.write get_entry_content(entry, data_provider) if entry.ftype != :directory
  end
  
  def replace_entry_content(str, data_provider)
    possible_tags = str.scan(all_tags_regex)
    # Loops through what looks like are tags. Anything with ||name|| even if they are not in the available tags list
    possible_tags.each do |tag|
      tag_name = extract_tag_name(tag)
      tag_name = squish_tag_name(tag, tag_name)
      tag_name = tag_name.to_s.to_sym
      # if in the available tag list, replace with the new value
      if data_provider.has_key?(tag_name)
        encoder = HTMLEntities.new
        content = encoder.encode("#{data_provider[tag_name]}")
        str.gsub!(tag, content)
      end
    end
    str
  end
  
  # extracts just the tag name
  def extract_tag_name(tag)
    malformed_tag_regex.match(tag)
    tag_name ||= well_formed_tag_regex.match(tag)
    tag_name ||= ''
  end
  
  # This will handle most instances where someone edits just part of a tag and Word wraps that part in more XML
  # If the tag did not have any extra xml formatting we just return the passed in tag_name
  def squish_tag_name(tag, tag_name)
    words = tag.scan(just_label_regex).flatten!
    if words.respond_to?(:size) && words.size > 1
      #Then the tag was split by word
      tag_name = words.join('')
    end
    tag_name
  end
  
end