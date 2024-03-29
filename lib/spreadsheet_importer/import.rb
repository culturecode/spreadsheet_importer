module SpreadsheetImporter
  module Import
    def self.from_xlsx(file_path, options = {}, &block)
      doc = Roo::Excelx.new(file_path, :file_warning => :ignore)

      options = {:sheet_name => nil}.merge(options)
      options[:start_row] -= doc.first_row - 1 if options[:start_row] # Roo starts at the first blank row, so compensate

      spreadsheet = []
      doc.each_with_pagename do |name, sheet|
        spreadsheet.concat sheet.to_a unless options[:sheet_name] && name.downcase.strip != options[:sheet_name].downcase.strip
      end

      raise MissingRequiredSheet, "Spreadsheet must include a Sheet named '#{options[:sheet_name]}'" if options[:sheet_name] && spreadsheet.empty?

      from_spreadsheet(spreadsheet, options, &block)
    end

    def self.from_csv(file_path, options = {}, &block)
      # Detect the encoding of the file and normalize it to UTF-8
      csv = File.read(file_path)
      encoding = CharlockHolmes::EncodingDetector.detect(csv)[:encoding]
      csv = CharlockHolmes::Converter.convert csv, encoding, 'UTF-8'

      # Get rid of the UTF-8 BOM since Charlock doesn't do this for us
      csv.slice!(0) if csv[0].ord == 65279

      # Determine whether the column separator is a tab or comma
      col_sep = csv.count("\t") > 0 ? "\t" : ","

      spreadsheet = CSV.parse(csv, :col_sep => col_sep, :headers => true, :header_converters => :downcase).to_a
      from_spreadsheet(spreadsheet, options, &block)
    end

    # Returns 2D array of the spreadsheet, rows by columns
    # If a block is given, yields each row to the block
    # Exceptions during the iteration will collected along with their row number
    # and re-raised at the end of processing
    def self.from_spreadsheet(spreadsheet, options = {}, &block)
      options = {:start_row => 1, :schema => nil}.merge(options)

      # Remove intro rows
      (options[:start_row] - 1).times { spreadsheet.shift }

      headers = spreadsheet.first || []
      if options[:required_columns]
        assert_required_columns!(headers, options[:required_columns])
      end

      if options[:schema] # If a Conformist schema is provided, use that to prepare rows
        rows = options[:schema].conform(spreadsheet, :skip_first => true)
      else
        spreadsheet.shift # Don't iterate over the headers
        rows = spreadsheet
      end

      # Create an enumerator the standardizes the output from conformist or 2D array spreadsheet
      errors = []
      spreadsheet = Spreadsheet.new do |yielder|
        row_number = options[:start_row]
        rows.each_with_index do |row, index|
          begin
            yielder.yield row_to_attributes(row, headers), index, row_number
          rescue => e
            error = "Row #{row_number}: #{e.message}"
            error << "\n#{e.backtrace.join("\n")}" if options[:error_backtrace]
            errors << error
          end
          row_number += 1
        end
      end
      spreadsheet.headers = headers
      spreadsheet.errors = errors
      spreadsheet.each(&block) if block_given?

      return spreadsheet
    end

    private

    def self.assert_required_columns!(headers, required_columns)
      required_columns.each do |column_name|
        raise MissingRequiredColumn, "Spreadsheet must include a '#{column_name}' column" unless column_present?(headers, column_name)
      end
    end

    def self.column_present?(headers, column_name)
      headers.collect {|c| c.to_s.downcase.squish }.include?(column_name.downcase.squish)
    end

    # Returns an attributes hash for the given row
    def self.row_to_attributes(row, headers)
      case row
      when Array
        Hash[headers.zip(row)]
      else # Handle Conformist::HashStruct rows
        row.attributes
      end
    end

  end

  # Spreadsheet

  class Spreadsheet < Enumerator
    attr_accessor :errors, :headers
  end

  # EXCEPTIONS

  class MissingRequiredSheet < StandardError; end
  class MissingRequiredColumn < StandardError; end
end
