module SpreadsheetImporter
  module Import
    def self.from_xlsx(file_path, options = {}, &block)
      options = {:sheet_name => nil}.merge(options)

      spreadsheet = []
      Roo::Excelx.new(file_path, :file_warning => :ignore).each_with_pagename do |name, sheet|
        spreadsheet.concat sheet.to_a unless options[:sheet_name] && name.downcase.strip != options[:sheet_name].downcase.strip
      end
      from_spreadsheet(spreadsheet, options, &block)
    end

    def self.from_csv(file_path, options = {}, &block)
      # Detect the encoding of the file and normalize it to UTF-8
      csv = File.read(file_path)
      encoding = CharlockHolmes::EncodingDetector.detect(csv)[:encoding]
      csv = CharlockHolmes::Converter.convert csv, encoding, 'UTF-8'

      # Get rid of the UTF-16LE BOM since Charlock doesn't do this for us
      csv.slice!(0) if csv[0].ord == 65279

      # Determine whether the column separator is a tab or comma
      col_sep = csv.count("\t") > 0 ? "\t" : ","

      spreadsheet = CSV.parse(csv, :col_sep => col_sep, :headers => true, :header_converters => :downcase)
      from_spreadsheet(spreadsheet, options, &block)
    end

    def self.from_spreadsheet(spreadsheet, options = {}, &block)
      options = {:start_row => 1, :schema => nil}.merge(options)

      (options[:start_row] - 1).times { spreadsheet.shift } # Remove intro rows
      spreadsheet = options[:schema].conform(spreadsheet) if options[:schema] # If a Conformist schema is provided, use that to prepare rows

      errors = []
      rowcount = 0
      spreadsheet.each do |row|
        rowcount += 1
        begin
          block.call(row)
          print '.'
        rescue => e
          progress_indicator = '!'
          errors << "Row #{rowcount}: #{e.message}"
          print '!'
        end
      end

      rowcount += options[:start_row] if rowcount > 0

      return {:imported => rowcount - errors.count, :errors => errors, :total => rowcount}
    end
  end
end
