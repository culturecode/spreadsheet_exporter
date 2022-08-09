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
  {"name" => "Jim", "role" => "admin", "city" => CITIES.sample},
  {"name" => "Sally", "role" => "user"},
  {"name" => "Horatio", "role" => "user", "meal" => "Paleo"},
  {"name" => "Jan", "role" => "user"}
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
      "source" => CITIES
    },
    "meal" => {
      "ignore_blank" => true,
      "error_type" => "warning",
      "source" => %w[Omnivore Veg Vegan]
    }
  }
}

File.binwrite("output.xlsx", SpreadsheetExporter::XLSX.from_objects(data, options))
