require 'roo'

module SpreadsheetImporter
  # Historically this gem used culturecode-roo, a fork whose only behavioural changes were:
  #
  #   1. "General" format cells are read as strings instead of numbers, so values like
  #      identification numbers round-trip exactly as typed (e.g. "1234", not 1234.0).
  #   2. Iterating an empty sheet yields nothing instead of raising.
  #
  # The gem now uses stock roo. This preserves behaviour 1 by treating the "General"
  # number format as a string type. Behaviour 2 is handled in Import.from_xlsx.
  module RooGeneralFormatAsString
    def _to_type(format)
      format.to_s.downcase == 'general' ? :string : super
    end
  end
end

Roo::Excelx::Format.singleton_class.prepend(SpreadsheetImporter::RooGeneralFormatAsString)
