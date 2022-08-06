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
  {"name" => "Jim", "role" => "admin", "city" => "Vancouver"},
  {"name" => "Sally", "role" => "user"},
  {"name" => "Horatio", "role" => "user", "meal" => "Paleo"},
  {"name" => "Jan", "role" => "user", "site_type" => SITE_TYPES.sample}
]

options = {
  "validations" => {
    "role" => {
      "ignore_blank" => false,
      "source" => %w[admin user spammer boss]
    },
    "city" => {
      "ignore_blank" => true,
      "error_type" => "information",
      "source" => %w[Victoria Vancouver Courtenay]
    },
    "meal" => {
      "ignore_blank" => true,
      "error_type" => "warning",
      "source" => %w[Omnivore Veg Vegan]
    },
    "site_type" => {
      "source" => SITE_TYPES
    }
  }
}

File.open("output.xlsx", "wb") do |f|
  f.write SpreadsheetExporter::XLSX.from_objects(data, options)
end
