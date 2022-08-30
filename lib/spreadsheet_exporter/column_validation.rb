module SpreadsheetExporter
  ColumnValidation = Struct.new(:ignore_blank, :data_source, :dependent_on, :error_type, keyword_init: true) do
    def initialize(*)
      super
      self.ignore_blank = true if ignore_blank.nil?
      self.error_type ||= VALIDATION_ERROR_TYPES[0]
    end
  end
end
