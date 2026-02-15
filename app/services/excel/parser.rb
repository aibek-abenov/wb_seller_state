require "roo"

module Excel
  class Parser
    # Roo индексы — с нуля
    COLUMN_F_INDEX = 5  # Название 1
    COLUMN_G_INDEX = 6  # Название 2
    COLUMN_I_INDEX = 8  # Баркод
    COLUMN_K_INDEX = 10 # Обоснование для оплаты

    SALE_REASON = "Продажа".freeze

    def self.call(file:)
      new(file).call
    end

    def initialize(file)
      @xlsx = Roo::Excelx.new(file.path)
    end

    def call
      rows = extract_rows

      rows = rows.select { |r| r[:sku].present? }
      rows = rows.select { |r| r[:payment_reason] == SALE_REASON }

      # Убираем дубли по баркоду
      rows.uniq { |r| r[:sku] }.map do |r|
        {
          sku: r[:sku],
          name_primary: r[:name_primary],
          name_secondary: r[:name_secondary]
        }
      end
    end

    private

    def extract_rows
      @xlsx.each_row_streaming(offset: 1).map do |row|
        {
          sku: fetch_cell(row, COLUMN_I_INDEX),
          name_primary: fetch_cell(row, COLUMN_F_INDEX),
          name_secondary: fetch_cell(row, COLUMN_G_INDEX),
          payment_reason: fetch_cell(row, COLUMN_K_INDEX)
        }
      end
    end

    def fetch_cell(row, index)
      cell = row[index]
      cell&.value.to_s.strip
    end
  end
end
