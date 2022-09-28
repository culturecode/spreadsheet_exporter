#!/usr/bin/env ruby
require_relative "../lib/spreadsheet_exporter"
require_relative "./fixtures"

data = [
  {"name" => "Jim", "role" => "admin"}.merge(country_and_city),
  {"name" => "Sally", "role" => "user", "favourite_meal" => MEALS.sample, "most_recent_meal" => MEALS.sample}.merge(country_and_city),
  {"name" => "Horatio", "role" => "user", "favourite_meal" => MEALS.sample, "most_recent_meal" => MEALS.sample},
  {"name" => "Jan", "role" => "user"}
]

options = {
  :data_sources => {
    "all_meals" => MEALS,
    "roles" => %w[admin user spammer boss],
    "countries" => COUNTRIES,
    "cities" => CONDITIONAL_CITIES
  },

  :validations => {
    "role" => SpreadsheetExporter::ColumnValidation.new(
      ignore_blank: false,
      data_source: "roles"
    ),
    "country" => SpreadsheetExporter::ColumnValidation.new(
      ignore_blank: true,
      error_type: "information",
      data_source: "countries"
    ),
    "city" => SpreadsheetExporter::ColumnValidation.new(
      ignore_blank: true,
      error_type: "information",
      dependent_on: "country",
      data_source: "cities"
    ),
    "favourite_meal" => SpreadsheetExporter::ColumnValidation.new(
      ignore_blank: true,
      error_type: "warning",
      data_source: "all_meals"
    ),
    "most_recent_meal" => SpreadsheetExporter::ColumnValidation.new(
      ignore_blank: true,
      error_type: "warning",
      data_source: "all_meals"
    )
  }
}
class Humanizer
  def self.human_attribute_name(att)
    att.upcase
  end
end

options[:humanize_headers_class] = Humanizer
options[:freeze_panes] = [1, 1]

File.binwrite("output.xlsx", SpreadsheetExporter::XLSX.from_objects(data, **options))
