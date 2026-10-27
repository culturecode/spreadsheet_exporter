require 'minitest/autorun'
require 'date'
require 'tmpdir'
require 'roo'
require 'spreadsheet_exporter/xlsx'

class XLSXTest < Minitest::Test
  def test_preserves_leading_zeros_in_numeric_strings
    with_export('file_number' => '0012345') do |spreadsheet|
      assert_equal '0012345', spreadsheet.cell(2, 1)
    end
  end

  def test_preserves_other_identifiers_as_text
    with_export('file_number' => '4007116', 'long_file_number' => '12345678901234567', 'postal_code' => '00123', 'reference' => 'A-12/34') do |spreadsheet|
      assert_equal ['4007116', '12345678901234567', '00123', 'A-12/34'], spreadsheet.row(2)
    end
  end

  def test_preserves_numeric_values_as_numbers
    with_export('id' => 12345, 'amount' => 12.5) do |spreadsheet|
      assert_equal [12345, 12.5], spreadsheet.row(2)
      assert_equal :float, spreadsheet.celltype(2, 1)
      assert_equal :float, spreadsheet.celltype(2, 2)
    end
  end

  def test_preserves_numeric_headers_as_text
    with_export('0012345' => 'value') do |spreadsheet|
      assert_equal '0012345', spreadsheet.cell(1, 1)
    end
  end

  def test_preserves_boolean_text_and_blank_values
    with_export('enabled' => true, 'disabled' => false, 'missing' => nil, 'empty' => '') do |spreadsheet|
      assert_equal 'true', spreadsheet.cell(2, 1)
      assert_equal 'false', spreadsheet.cell(2, 2)
      assert_equal '', spreadsheet.cell(2, 3).to_s
      assert_equal '', spreadsheet.cell(2, 4).to_s
    end
  end

  def test_preserves_date_and_array_output
    date = Date.new(2026, 10, 27)
    time = Time.utc(2026, 10, 27, 12)
    datetime = DateTime.new(2026, 10, 27, 12)

    with_export('date' => date, 'time' => time, 'datetime' => datetime, 'references' => ['0012345', 42, false, nil]) do |spreadsheet|
      assert_equal [date.to_s, time.to_s, datetime.to_s, '0012345, 42, false, '], spreadsheet.row(2)
    end
  end

  private

  def with_export(values)
    record = Struct.new(:values) do
      def as_csv(_options = {})
        values
      end
    end.new(values)

    Dir.mktmpdir('spreadsheet-exporter-test') do |directory|
      Dir.chdir(directory) do
        Dir.mkdir('tmp')
        File.binwrite('export.xlsx', SpreadsheetExporter::XLSX.from_objects([record]))
        spreadsheet = Roo::Excelx.new('export.xlsx')

        yield spreadsheet
      ensure
        spreadsheet&.close
      end
    end
  end
end
