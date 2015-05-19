# TODO: Find out why we can't detect arrays properly and must resort to crappy class.name comparison
module SpreadsheetExporter
  module Spreadsheet
    def self.from_objects(objects, options = {})
      headers = []
      rows = []

      # Get all the data and accumulate headers from each row (since rows may not have all the same attributes)
      Array(objects).each do |object|
        data = object.respond_to?(:as_csv) ? get_values(object.as_csv(options)) : get_values(object.as_json(options))
        headers = headers | data.keys
        rows << data
      end

      # Create the csv, ensuring to place each row's attributes under the appropriate header (since rows may not have all the same attributes)
      [].tap do |spreadsheet|
        spreadsheet << headers
        rows.each do |row|
          sorted_row = []
          row.each do |header, value|
            sorted_row[headers.index(header)] = value
          end

          spreadsheet << sorted_row
        end
      end
    end

    # Return an array of human_attribute_name's
    # Used by the CSV Import/Export process to match CSV headers to model attribute names
    def self.han(klass, *attributes)
      options = attributes.extract_options!

      attributes.flatten!
      attributes.collect! {|attribute| klass.human_attribute_name(attribute) }
      attributes.collect!(&:downcase) if options[:downcase]

      return attributes.many? ? attributes : attributes.first
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
