# SpreadsheetExporter

```ruby
gem 'spreadsheet_exporter'
```


## Usage

Objects that are exported must respond to `as_spreadsheet`, `as_csv` or `as_json` and return a hash
representing column names and values.  In Rails you can `include SpreadsheetExporter::GeneratesSpreadsheet` into your model.


### CSV or XLSX
Output can be .csv or .xlsx. Choose by using SpreadsheetExporter::CSV or SpreadsheetExporter::XLSX modules.
Note that .csv output is actually tab delimited by default so that Excel can
open files properly without needing to use its import function. If you need to output
that is actually comma-delimited, pass ```:col_sep => ','``` as an option when exporting

### Array of ActiveRecord Objects
```ruby
  SpreadsheetExporter::CSV.from_objects(array_of_objects, options)

  # Humanize header names using klass.human_attribute_name
  SpreadsheetExporter::CSV.from_objects(array_of_objects, :humanize_headers_class => User)
```

### 2D Array
```ruby
  SpreadsheetExporter::CSV.from_spreadsheet([["First Name", "Last Name"], ["Bob", "Hoskins"], ["Roger", "Rabbit"]])
```

### XLSX with Pick Lists

```ruby
# data sources are written to a `data` worksheet and may be referenced by
# multiple rows
data_sources = {
  "food_types" => %w[Polenta Paella Papaya],
  "countries" => %w[Canada Türkiye],
  "cities" => {"Canada"=>["Sxwōxwiyám", "Toronto"], "Türkiye"=>["Eskişehir", "İzmir", "İstanbul"]}
}

validations = {
  "favourite_food" => SpreadsheetExporter::ColumnValidation.new(
    data_source: "food_types"
  ),
  "yuckiest_food" => SpreadsheetExporter::ColumnValidation.new(
    data_source: "food_types"
  ),
  "country" => SpreadsheetExporter::ColumnValidation.new(
    data_source: "countries"
  ),
  "city" => SpreadsheetExporter::ColumnValidation.new(
    dependent_on: "country",
    data_source: "cities"
  ),
}

SpreadsheetExporter::XLSX.from_objects(array_of_objects,
  data_sources: data_sources,
  validations: validations,
  freeze_panes: [1, 0] # number of rows and columns to freeze (only applies to XLSX)
)
```

### Testing

There isn't currently a test suite.  You can generate a sample test XLSX file by running `test/test.rb`.
