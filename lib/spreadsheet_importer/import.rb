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

    # Returns 2D array of the spreadsheet, rows by columns
    # If a block is given, yields each row to the block
    # Exceptions during the iteration will collected along with their row number
    # and re-raised at the end of processing
    def self.from_spreadsheet(spreadsheet, options = {}, &block)
      options = {:start_row => 1, :schema => nil}.merge(options)

      # Remove intro rows
      (options[:start_row] - 1).times { spreadsheet.shift }

      headers = spreadsheet.first

      if options[:schema] # If a Conformist schema is provided, use that to prepare rows
        spreadsheet = options[:schema].conform(spreadsheet, :skip_first => true)
      else
        spreadsheet = spreadsheet
      end

      # Make it possible to add errors to the spreadsheet
      class << spreadsheet
        def errors=(array)
          @errors = array
        end

        def errors
          @errors ||= []
        end
      end


      if block_given?
        row_number = options[:start_row]
        spreadsheet.each_with_index do |row, index|
          begin
            block.call(row_to_attributes(row, headers), index, row_number)
            print '.'
          rescue => e
            spreadsheet.errors << "Row #{row_number}: #{e.message}\n#{e.backtrace.join("\n")}"
            print '!'
          end
          row_number += 1
        end
      end

      return spreadsheet
    end

    private

    # Returns an attributes hash for the given row
    def self.row_to_attributes(row, headers)
      case row
      when Array
        Hash[ItemSchema.attribute_names(headers).zip(row)]
      else # Handle Conformist::HashStruct rows
        row.attributes
      end
    end

  end
end
