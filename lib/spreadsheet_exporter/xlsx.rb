require 'write_xlsx'
require_relative 'spreadsheet'

module SpreadsheetExporter
  module XLSX
    extend Writexlsx::Utility # gets us `xl_rowcol_to_cell`

    ROW_MAX = 65_536 - 1

    # Excel allows defining validation `sources` in two different ways, an inline
    # list or a reference to cells elsewhere in the workbook.
    #
    # The inline list is defined as a comma-separated string with a max length of
    # 255 characters.
    USE_INLINE_LISTS = false # debug toggle, not for production
    MAX_INLINE_LIST_CHARS = 255

    VALIDATION_ERROR_TYPES = %w[stop warning information].freeze
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

      add_worksheet_validation(workbook, worksheet, column_indexes, header_format, options)

      workbook.worksheets.each do |ws|
        ws.freeze_panes(1, 0)
      end

      workbook.close
      io.string
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

    def self.add_worksheet_validation(workbook, worksheet, column_indexes, header_format, options = {})
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

        validation_options = add_column_validation(workbook, column_name, column_index, column_validation, header_format)

        pp validation_options

        worksheet.data_validation(1, column_index, ROW_MAX, column_index, validation_options)
      end
    end

    def self.add_column_validation(workbook, column_name, column_index, column_validation, header_format)
      list_values = Array(column_validation.fetch("source", []))
      if list_values.empty?
        raise ArgumentError, "no values for validation for column '#{column_name}'"
      end

      error_type = column_validation.fetch("error_type", VALIDATION_ERROR_TYPES[0])
      unless VALIDATION_ERROR_TYPES.include?(error_type)
        raise ArgumentError, "invalid error_type `#{error_type}` for validation for column '#{column_name}'"
      end

      list_values.compact!
      list_length = list_values.join(",").length

      source = nil

      if USE_INLINE_LISTS && list_length <= MAX_INLINE_LIST_CHARS
        # commas are not allowed when
        # TODO: we should warn about losing any commas
        list_values.map! { |v| v.sub(',', '').strip }
        source = list_values
      else
        data_start = xl_rowcol_to_cell(1, column_index)
        data_end = xl_rowcol_to_cell(ROW_MAX, column_index)

        source = "=data!$#{data_start}:#{data_end}"
        warn "list values for column #{column_name} too long to be inlined, " \
          "len #{list_length} > #{MAX_INLINE_LIST_CHARS}, moving source to #{source}"

        unless (data_sheet = workbook.worksheet_by_name(DATA_WORKSHEET_NAME))
          data_sheet = workbook.add_worksheet(DATA_WORKSHEET_NAME)
        end

        data_sheet.write(0, column_index, column_name, header_format)
        data_sheet.write_col(1, column_index, list_values)
      end

      {
        "validate" => "list",
        "input_title" => "Select a value",
        "source" => source,
        "error_message" => column_validation.fetch("error_message", "Please select a valid option"),
        "error_type" => error_type,
        "ignore_blank" => column_validation.fetch("ignore_blank", true),
        "dropdown" => true
      }
    end
  end
end
