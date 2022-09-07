require 'write_xlsx'
require_relative 'spreadsheet'
require "active_support"
require "active_support/core_ext/hash/keys"

module SpreadsheetExporter
  module XLSX
    extend Writexlsx::Utility # gets us `xl_rowcol_to_cell` and `xl_col_to_name`

    ROW_MAX = 65_536 - 1
    DATA_WORKSHEET_NAME = "data".freeze

    def self.from_objects(objects, humanize_headers_class: nil, **options)
      spreadsheet = Spreadsheet.from_objects(objects, humanize_headers_class: humanize_headers_class, **options).compact
      from_spreadsheet(spreadsheet, **options)
    end

    def self.from_spreadsheet(spreadsheet, validations: {}, data_sources: {}, freeze_panes: false, **options)
      io = StringIO.new
      workbook = WriteXLSX.new(io)

      worksheet = workbook.add_worksheet

      header_format = workbook.add_format
      header_format.set_bold

      column_indexes = {}

      # Write header row
      Array(spreadsheet.first).each_with_index do |header, col|
        worksheet.write(0, col, header.to_s, header_format)
        column_indexes[header.attribute_name] = col
      end

      Array(spreadsheet[1..]).each_with_index do |values, row|
        worksheet.write_row(row + 1, 0, Array(values))
      end

      added_data_sources = add_data_sources(workbook, header_format, data_sources)

      add_worksheet_validation(workbook, worksheet, column_indexes, added_data_sources, header_format, validations)

      add_frozen_panes(worksheet, freeze_panes)

      workbook.close
      io.string
    end

    def self.sanitize_defined_name(raw)
      raw.gsub(/[^A-Za-z0-9_]/, "_")
    end

    # freeze_panes => [1, 2] # freeze the top row and left two cols
    def self.add_frozen_panes(worksheet, freeze_panes)
      return unless freeze_panes
      rows, cols = freeze_panes
      worksheet.freeze_panes(Integer(rows), Integer(cols))
    end

    # Write each data_source to the `data` worksheet and wrap it with a named range so
    # we can easily reference it later.
    #
    # `data_sources` is a hash in the format:
    #     { 'data_source_id' => ['data', 'source', 'options'] }
    #
    # This will create a named range called `data_source_id`.
    #
    # For data sources dependent on the value in another column, the format is
    #     { 'data_source_id' => {
    #         'other_col_val_1' => ['options', 'when', 'val is val_1'],
    #         'other_col_val_2' => ['options', 'when', 'val is val_2']
    #       }
    #     }
    #
    # This will create two named ranges: `data_source_id_val_1` and `data_source_id_val_2`.
    def self.add_data_sources(workbook, header_format, data_sources)
      return {} if data_sources.empty?

      unless (data_sheet = workbook.worksheet_by_name(DATA_WORKSHEET_NAME))
        data_sheet = workbook.add_worksheet(DATA_WORKSHEET_NAME)
        data_sheet.freeze_panes(1, 0)
      end

      data_source_refs = {}

      column_index = 0
      data_sources.stringify_keys.each do |data_key, data_values|
        if data_values.is_a?(Hash)
          # this is a dependent data source
          data_values.each do |data_value, sub_values|
            sub_key = sanitize_defined_name("#{data_key}_#{data_value}")
            data_source_refs[sub_key] = add_data_source(workbook, data_sheet, sub_key, sub_values, column_index, header_format)
            column_index += 1
          end
        else
          # this is an independent data source
          data_source_refs[data_key] = add_data_source(workbook, data_sheet, data_key, data_values, column_index, header_format)
          column_index += 1
        end
      end

      data_source_refs
    end

    # Write a data column to the `data` worksheet and define it as a named range
    #
    # Returns the named range's name
    def self.add_data_source(workbook, data_sheet, data_key, data_values, column_index, header_format)
      unless data_values.is_a?(Array)
        raise ArgumentError, "data_values should be an array (got #{data_values.inspect}"
      end

      data_start = xl_rowcol_to_cell(1, column_index, true, true)
      data_end = xl_rowcol_to_cell(data_values.length, column_index, true, true)

      defined_name_source = "=#{DATA_WORKSHEET_NAME}!#{data_start}:#{data_end}"

      data_sheet.write(0, column_index, data_key, header_format)
      data_sheet.write_col(1, column_index, data_values.map(&:strip))
      workbook.define_name(data_key, defined_name_source)

      data_key
    end

    def self.add_worksheet_validation(workbook, worksheet, column_indexes, added_data_sources, header_format, validations)
      return if validations.empty?

      validations.each do |column_name, column_validation|
        column_index = column_indexes[column_name.to_s]

        if column_index.nil?
          warn "attempted to apply validation to missing column '#{column_name}'"
          next
        end

        defined_name = if column_validation.dependent_on
                         parent_col_index = column_indexes[column_validation.dependent_on]
                         parent_col = xl_col_to_name(parent_col_index, true)
                         dependent_named_range(column_validation.data_source, parent_col)
                       else
                         added_data_sources[column_validation.data_source]
                       end

        unless defined_name
          raise ArgumentError, "missing data for data_source=#{column_validation.data_source}, " \
                               "tried defined_name #{defined_name}"
        end

        validation_options = generate_validation(column_validation, defined_name)
        pp validation_options
        worksheet.data_validation(1, column_index, ROW_MAX, column_index, validation_options)
      rescue StandardError => e
        debugger
      end
    end

    # We build up the reference to the named range by leaning on Excel's INDIRECT function
    # to dynamically build the name.  The resulting formula becomes the validation drop down's
    # source. It resolves thusly...
    #
    # =INDIRECT("sub_data_source" & "_" & SUBSTITUTE(INDIRECT("$AA" & ROW()), " ", "_"))
    # =INDIRECT("sub_data_source" & "_" & SUBSTITUTE("Parent Value, " ", "_"))
    # =INDIRECT("sub_data_source" & "_" & "Parent_Value")
    # =INDIRECT("sub_data_source_Parent_Value")
    # =sub_data_source_Parent_Value
    def self.dependent_named_range(data_source, parent_col)
      "INDIRECT(\"#{data_source}\" & \"_\" & "\
      "SUBSTITUTE(INDIRECT(\"#{parent_col}\" & ROW()), \" \", \"_\"))"
    end

    def self.generate_validation(column_validation, defined_name)
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
