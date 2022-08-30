#!/usr/bin/env ruby
require_relative "./lib/spreadsheet_exporter"
require_relative "./test_data"
require "awesome_print"
require "debug"

# http://support.microsoft.com/kb/211485
#
# TODO: we should add a column for all the missing validations even if there is
# no data in it yet

data = [
  {"name" => "Jim", "role" => "admin"}.merge(country_and_city),
  {"name" => "Sally", "role" => "user", "favourite_meal" => MEALS.sample, "most_recent_meal" => MEALS.sample}.merge(country_and_city),
  {"name" => "Horatio", "role" => "user", "favourite_meal" => MEALS.sample, "most_recent_meal" => MEALS.sample},
  {"name" => "Jan", "role" => "user"}
]

options = {
  "data_sources" => {
    "all_meals" => MEALS,
    "roles" => %w[admin user spammer boss],
    "countries" => COUNTRIES,
    "cities" => CONDITIONAL_CITIES
  },

  "validations" => {
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
      indirect_built_from: "country",
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

File.binwrite("output.xlsx", SpreadsheetExporter::XLSX.from_objects(data, options))
