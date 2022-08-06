$:.push File.expand_path("../lib", __FILE__)

# Maintain your gem's version:
require "spreadsheet_exporter/version"

# Describe your gem and declare its dependencies:
Gem::Specification.new do |s|
  s.name        = "spreadsheet_exporter"
  s.version     = SpreadsheetExporter::VERSION
  s.authors     = ["Nicholas Jakobsen"]
  s.email       = ["nicholas.jakobsen@gmail.com"]
  s.homepage    = "https://github.com/culturecode/spreadsheet_exporter"
  s.summary     = "Export your data as various types of spreadsheets"
  s.description = "Export your data as various types of spreadsheets. Supports csv and xlsx output."
  s.license     = "MIT"

  s.files = Dir["{app,config,db,lib}/**/*", "MIT-LICENSE", "Rakefile", "README.md"]
  s.test_files = Dir["test/**/*"]

  s.add_dependency "activesupport", ">= 6"
  s.add_dependency "write_xlsx"
end
