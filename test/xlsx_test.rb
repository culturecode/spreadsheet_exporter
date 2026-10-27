require 'minitest/autorun'
require 'tmpdir'
require 'roo'
require 'spreadsheet_exporter/xlsx'

class XLSXTest < Minitest::Test
  def test_preserves_leading_zeros_in_numeric_strings
    record = Object.new
    def record.as_csv(_options = {})
      { 'file_number' => '0012345' }
    end

    Dir.mktmpdir('spreadsheet-exporter-test') do |directory|
      Dir.chdir(directory) do
        Dir.mkdir('tmp')
        File.binwrite('export.xlsx', SpreadsheetExporter::XLSX.from_objects([record]))
        spreadsheet = Roo::Excelx.new('export.xlsx')

        assert_equal '0012345', spreadsheet.cell(2, 1)
      ensure
        spreadsheet&.close
      end
    end
  end
end
