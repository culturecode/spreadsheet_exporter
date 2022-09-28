# TODO: Find out why we can't detect arrays properly and must resort to crappy class.name comparison
module SpreadsheetExporter
  HeaderCell = Struct.new(:attribute_name, :human_attribute_name) do
    def to_s
      human_attribute_name.presence || attribute_name
    end
  end

  module Spreadsheet
    def self.from_objects(objects, humanize_headers_class: nil, **options)
      headers = []
      rows = []

      # Get all the data and accumulate headers from each row (since rows may not have all the same attributes)
      Array(objects).each do |object|
        data = if object.respond_to?(:as_spreadsheet)
          get_values(object.as_spreadsheet(options))
        elsif object.respond_to?(:as_csv)
          get_values(object.as_csv(options))
        else
          get_values(object.as_json(options))
        end

        headers |= data.keys.map { |v| HeaderCell.new(v) }
        rows << data
      end

      # Create the csv, ensuring to place each row's attributes under the appropriate header (since rows may not have all the same attributes)
      [].tap do |spreadsheet|
        if humanize_headers_class
          headers = han(headers, humanize_headers_class: humanize_headers_class, **options)
        end

        spreadsheet << headers

        rows.each do |row|
          sorted_row = []
          row.each do |header, value|
            col_index = headers.find_index { |h| h.attribute_name == header }
            sorted_row[col_index] = value
          end

          spreadsheet << sorted_row
        end
      end
    end

    # Return an array of human_attribute_name's
    # Used by the CSV Import/Export process to match CSV headers to model attribute names
    def self.han(headers, humanize_headers_class:, downcase: false, **)
      headers.flatten!

      headers.collect! do |header|
        header.human_attribute_name = humanize_headers_class.human_attribute_name(header.attribute_name)
        header.human_attribute_name.downcase! if downcase
        header
      end
    end

    def self.get_values(node, current_header = nil)
      output = {}
      case node.class.name
      when 'Hash'
        node.each do |key, subnode|
          output.merge! get_values(subnode, [current_header, key].compact.join('_'))
        end
      when 'Array'
        node.each do |subnode|
          get_values(subnode, current_header).each do |key, value|
            output[key] = [output[key], value.to_s].compact.join(', ')
          end
        end
      else
        output[current_header] = node.to_s
      end
      return output
    end
  end
end
