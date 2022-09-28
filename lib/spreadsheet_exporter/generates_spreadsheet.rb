require "active_support/concern"

module SpreadsheetExporter
  module GeneratesSpreadsheet
    extend ActiveSupport::Concern

    def as_spreadsheet(options = {})
      serializable_hash(options)
    end
  end
end
