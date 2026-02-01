module Excel
  class Parser
    COLUMN_F_INDEX = 5 # Название 1
    COLUMN_G_INDEX = 6 # Название 2
    COLUMN_I_INDEX = 8 # Артикул

    def self.call(file:)
      new(file).call
    end

    def initialize(file)
      @xlsx = Roo::Excelx.new(file.path)
    end

    def call
      rows = extract_rows

      # Убираем строки без артикула
      rows = rows.select { |r| r[:sku].present? }

      # Убираем дубли по артикулу
      rows.uniq { |r| r[:sku] }
    end

    private

    def extract_rows
      @xlsx.each_row_streaming(offset: 1).map do |row|
        {
          sku: fetch_cell(row, COLUMN_I_INDEX), # Артикул (ключ)
          name_primary: fetch_cell(row, COLUMN_F_INDEX), # Название (удобное для клиента)
          name_secondary: fetch_cell(row, COLUMN_G_INDEX) # Альтернативное название
        }
      end
    end

    def fetch_cell(row, index)
      cell = row[index]
      cell&.value.to_s.strip
    end
  end
end