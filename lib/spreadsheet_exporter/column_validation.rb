module SpreadsheetExporter
  ColumnValidation = Struct.new(:attribute_name, :ignore_blank, :data_source, :indirect_built_from, :error_type, keyword_init: true) do
    def initialize(*)
      super
      self.ignore_blank = true if ignore_blank.nil?
      self.error_type ||= VALIDATION_ERROR_TYPES[0]
    end
  end
end
