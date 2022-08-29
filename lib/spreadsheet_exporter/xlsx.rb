require 'write_xlsx'
require_relative 'spreadsheet'
require "active_support"
require "active_support/core_ext/hash/keys"

module SpreadsheetExporter
  module XLSX
    extend Writexlsx::Utility # gets us `xl_rowcol_to_cell`

    ROW_MAX = 65_536 - 1
    DATA_WORKSHEET_NAME = "data".freeze

    def self.from_objects(objects, options = {})
      spreadsheet = Spreadsheet.from_objects(objects, options).compact
      from_spreadsheet(spreadsheet, options.deep_stringify_keys)
    end

    def self.from_spreadsheet(spreadsheet, options = {})
      io = StringIO.new
      workbook = WriteXLSX.new(io)

      worksheet = workbook.add_worksheet

      header_format = workbook.add_format
      header_format.set_bold

      column_indexes = {}

      # Write header row
      Array(spreadsheet.first).each_with_index do |column_name, col|
        worksheet.write(0, col, column_name, header_format)
        column_indexes[column_name] = col
      end

      Array(spreadsheet[1..]).each_with_index do |values, row|
        worksheet.write_row(row + 1, 0, Array(values))
      end

      data_sources = add_data_sources(workbook, header_format, options)
      add_worksheet_validation(workbook, worksheet, column_indexes, data_sources, header_format, options)

      workbook.worksheets.each do |ws|
        ws.freeze_panes(1, 0)
      end

      workbook.close
      io.string
    end

    def self.add_data_sources(workbook, header_format, options = {})
      data_sources = options.fetch("data_sources", {}) || {}
      return {} if data_sources.empty?

      unless (data_sheet = workbook.worksheet_by_name(DATA_WORKSHEET_NAME))
        data_sheet = workbook.add_worksheet(DATA_WORKSHEET_NAME)
      end

      data_source_refs = {}

      data_sources.each_with_index do |(data_key, data_values), column_index|
        data_source_refs[data_key] = add_data_source(workbook, data_sheet, data_key, data_values, column_index, header_format)
      end

      data_source_refs
    end

    # Write a data column to the `data` worksheet and define it as a named range
    #
    # Returnd the named range's name
    def self.add_data_source(workbook, data_sheet, data_key, data_values, column_index, header_format)
      raise ArgumentError unless data_values.is_a?(Array)

      data_start = xl_rowcol_to_cell(1, column_index, true, true)
      data_end = xl_rowcol_to_cell(data_values.length, column_index, true, true)

      defined_name_source = "=#{DATA_WORKSHEET_NAME}!#{data_start}:#{data_end}"

      data_sheet.write(0, column_index, data_key, header_format)
      data_sheet.write_col(1, column_index, data_values)
      defined_name = data_key
      workbook.define_name(defined_name, defined_name_source)
      defined_name
    end

    # TODO: we should DRY this up with the Spreadsheet.from_objects logic
    def self.rewrite_validation_column_names(column_validations, options)
      return column_validations unless options["humanize_headers_class"]
      klass = options["humanize_headers_class"]

      column_validations.each_with_object({}) do |(attribute, v), obj|
        rewritten = klass.human_attribute_name(attribute)
        rewritten.downcase! if options[:downcase]
        obj[rewritten] = v
      end
    end

    def self.add_worksheet_validation(workbook, worksheet, column_indexes, data_sources, header_format, options = {})
      column_validations = options.fetch("validations", {}) || {}
      return if column_validations.empty?

      column_validations = rewrite_validation_column_names(column_validations, options)

      column_validations.each do |column_name, column_validation|
        column_index = column_indexes[column_name]

        if column_index.nil?
          # TODO: we should output an empty column anyways
          warn "attempted to apply validation to missing column '#{column_name}'"
          next
        end

        defined_name = data_sources[column_validation.data_source]
        raise ArgumentError, "missing data for data_source=#{column_validation.data_source}" unless defined_name

        validation_options = add_column_validation(column_validation, defined_name)

        pp validation_options

        worksheet.data_validation(1, column_index, ROW_MAX, column_index, validation_options)
      end
    end

    def self.add_column_validation(column_validation, defined_name)
      {
        "validate" => "list",
        "input_title" => "Select a value",
        "source" => "=#{defined_name}",
        "error_type" => column_validation.error_type,
        "ignore_blank" => column_validation.ignore_blank,
        "dropdown" => true
      }
    end
  end
end
