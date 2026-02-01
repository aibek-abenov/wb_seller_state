module Excel
  class FileChecker
    def self.call(file:)
      new(file).call
    end

    def initialize(file)
      @file = file
    end

    def call
      validate_extension!
      validate_size!

      file
    end

    private

    attr_reader :file

    def validate_extension!
      allowed = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ]

      raise "Неверный формат файла — загрузите .xlsx" unless allowed.include?(file.content_type)
    end

    def validate_size!
      raise "Файл слишком большой (макс. 10 МБ)" if file.size > 10.megabytes
    end
  end
end
