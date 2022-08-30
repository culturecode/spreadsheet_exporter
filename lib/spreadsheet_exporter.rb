require_relative './spreadsheet_exporter/column_validation'
require_relative './spreadsheet_exporter/csv'
require_relative './spreadsheet_exporter/xlsx'
require 'active_support'
require 'active_support/core_ext/object/json'
require 'active_support/core_ext/hash/reverse_merge'

module SpreadsheetExporter
  VALIDATION_ERROR_TYPES = %w[stop warning information].freeze

  begin
    Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx
  rescue NameError
    # No Mime::Type
  end
end
