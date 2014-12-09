[![Build Status](https://travis-ci.org/mmmries/docx_templater.png)](https://travis-ci.org/mmmries/docx_templater)
[![Code Climate](https://codeclimate.com/github/mmmries/docx_templater.png)](https://codeclimate.com/github/mmmries/docx_templater)
[![Dependency Status](https://gemnasium.com/mmmries/docx_templater.png)](https://gemnasium.com/mmmries/docx_templater)

Docx Templater
==============

Use .docx files as reusable templates

In your word document put placeholder names in double pipes || e.g. `||client_email1||`

Example usage:
```ruby
  buffer = DocxTemplater.new.replace_file_with_content('path/to/your.docx',
     {
       :client_email1 => 'test@example.com',
       :client_phone1 => '555-555-5555',
     })

  # In Rails you can send a docx file via send_data
  send_data fake_buffer, :filename => 'REPC.docx'

  # Or save the output to a file
  File.open("path/to/mydocument.docx", "wb") {|f| f.write(buffer.string) }
```

Newline Conversion
==================

By default newlines in replacement values are converted to docx linebreaks. If you need to ignore newlines you can pass a flag.

```ruby
DocxTemplater.new(convert_newlines: false)
  .replace_file_with_content('spec/fixtures/newlines.docx',
    {quotes: "Be Excellent\nTo each other.\n~Bill and Ted's"})
```

Planned Changes
===============

* support for different tags (ie replace &lt;&lt;foo&gt;&gt; instead of ||foo||)
