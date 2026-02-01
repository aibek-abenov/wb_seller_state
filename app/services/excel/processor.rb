require "roo"
require "caxlsx"

module Excel
  class Processor
    # индексы Roo — с нуля
    SELECTED_COLS = {
      f:  5,
      i:  8,

      ah: 33,
      ai: 34,
      aj: 35,
      ak: 36,
      ao: 40,
      aq: 42,

      bh: 59,
      bi: 60,
      bj: 61
    }.freeze

    DO_NOT_SUM = %i[f i].freeze

    def self.call(file_path:, pricing:)
      new(file_path, pricing).call
    end

    def initialize(file_path, pricing)
      @file_path = file_path
      @pricing   = pricing
      @pricing_by_sku = pricing.index_by { |p| p[:sku].to_s }
    end

    def call
      workbook = Roo::Excelx.new(@file_path)
      sheet = workbook.sheet(0)

      rows = []
      sheet.each_with_index do |row, idx|
        next if idx == 0 # header
        rows << extract_selected_columns(row)
      end

      result_rows = []
      header = SELECTED_COLS.keys.map { |k| k.to_s.upcase } + ["PURCHASE_PRICE", "EXTRA_COSTS"]
      result_rows << header

      rows.each_slice(2) do |row1, row2|
        merged = merge_pair(row1, row2)
        sku = merged[:i].to_s

        price = @pricing_by_sku[sku]

        result_rows << merged.values + [
          price&.dig(:purchase_price),
          price&.dig(:extra_costs)
        ]
      end

      result_rows = add_profits_and_margin_percentages(result_rows)
      result_row_totals = calculate_totals(result_rows)

      result_rows.concat(result_row_totals)

      export_excel(result_rows)
    end

    private

    def extract_selected_columns(row)
      SELECTED_COLS.transform_values { |idx| row[idx] }
    end

    def merge_pair(a, b)
      a.each_with_object({}) do |(key, v1), acc|
        v2 = b[key]

        if DO_NOT_SUM.include?(key)
          acc[key] = v1
        elsif numeric?(v1) && numeric?(v2)
          acc[key] = v1.to_f + v2.to_f
        else
          acc[key] = [v1, v2].compact.map { |v| v.to_s.strip }.reject(&:empty?).join(" ")
        end
      end
    end

    def numeric?(value)
      return true if value.is_a?(Numeric)
      Float(value)
      true
    rescue StandardError
      false
    end

    # Безопасно превращаем в число (nil/строка/число) -> Float
    def n(value)
      return 0.0 if value.nil?
      return value.to_f if value.is_a?(Numeric)
      Float(value)
    rescue StandardError
      value.to_f
    end

    # result_rows уже включает:
    # 0: F
    # 1: I
    # 2: AH
    # 3: AI
    # 4: AJ
    # 5: AK
    # 6: AO
    # 7: AQ
    # 8: BH
    # 9: BI
    # 10: BJ
    # 11: PURCHASE_PRICE
    # 12: EXTRA_COSTS
    def add_profits_and_margin_percentages(result_rows)
      # header
      result_rows[0] = result_rows[0] + ["NET_PROFIT", "MARGIN_PERCENTAGE"]

      result_rows.each_with_index do |row, idx|
        next if idx == 0 # пропускаем header

        ah = n(row[2])

        net_profit =
          if ah.positive?
            ah - (
              n(row[5]) +  # AK
                n(row[6]) +  # AO
                n(row[8]) +  # BH
                n(row[9]) +  # BI
                n(row[10]) + # BJ
                n(row[11]) + # PURCHASE_PRICE
                n(row[12])   # EXTRA_COSTS
            )
          else
            0.0
          end

        denominator = n(row[11]) + n(row[12]) # purchase_price + extra_costs

        margin_percentage =
          if net_profit.positive? && denominator.positive?
            (net_profit * 100.0) / denominator
          else
            0.0
          end

        # дописываем в текущую строку (in-place)
        row.concat([net_profit, margin_percentage.round(2)])
      end

      result_rows
    end

    def calculate_totals(result_rows)
      titles = [
        "Общее к перечислению продавцу",     # AH (2)
        "Общее кол-во доставок",             # AI (3)
        "Общее кол-во возврата",             # AJ (4)
        "Общая сумма доставок",              # AK (5)
        "Общая сумма штрафов",               # AO (6)
        "Общая сумма хранения",              # BH (8)
        "Общая сумма удержании",             # BI (9)
        "Общая сумма платных приемок",       # BJ (10)
        "Общая сумма закупки",               # PURCHASE_PRICE (11)
        "Общая сумма прочих расходов",       # EXTRA_COSTS (12)
        "Общая сумма чистой прибыли",        # NET_PROFIT (13)
        "Общий процент маржинальности"
      ]

      col_indexes = [2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13]

      totals = Array.new(col_indexes.length, 0.0)

      result_rows.each_with_index do |row, idx|
        next if idx == 0 # header

        col_indexes.each_with_index do |col_idx, j|
          totals[j] += n(row[col_idx])
        end
      end

      # totals indexes:
      # 8  => PURCHASE_PRICE sum
      # 9  => EXTRA_COSTS sum
      # 10 => NET_PROFIT sum
      total_purchase_price = totals[8]
      total_extra_costs    = totals[9]
      total_net_profit     = totals[10]

      denom = total_purchase_price + total_extra_costs

      overall_margin_percentage =
        if total_net_profit.positive? && denom.positive?
          (total_net_profit * 100.0) / denom
        else
          0.0
        end

      [titles, totals + [overall_margin_percentage.round(2)]]
    end


    def export_excel(rows)
      path = "/tmp/result_#{Time.now.to_i}.xlsx"

      Axlsx::Package.new do |p|
        p.workbook.add_worksheet(name: "Result") do |sheet|
          rows.each { |r| sheet.add_row(r) }
        end

        p.serialize(path)
      end

      path
    end
  end
end
