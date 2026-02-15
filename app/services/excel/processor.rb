require "roo"
require "caxlsx"

module Excel
  class Processor
    SALE_REASON = "Продажа".freeze

    # Roo индексы — с нуля
    SELECTED_COLS = {
      supplier_article:      5,   # F  — Артикул поставщика
      barcode:               8,   # I  — Баркод
      payment_reason:        10,  # K  — Обоснование для оплаты

      payout_amount:         33,  # AH — К перечислению продавцу...
      delivery_service_fee:  36   # AK — Логистика
    }.freeze

    def self.call(file_path:, pricing:)
      new(file_path, pricing).call
    end

    def initialize(file_path, pricing)
      @file_path = file_path
      @pricing_by_sku = pricing.index_by { |p| p[:sku].to_s }
    end

    # Возвращает:
    # {
    #   path: "/tmp/xxx.xlsx",
    #   totals: { titles: [...], values: [...] }
    # }
    def call
      workbook = Roo::Excelx.new(@file_path)
      sheet    = workbook.sheet(0)

      header = [
        "Barcode",
        "SupplierArticle",
        "PaymentReason",
        "PayoutAmount",
        "DeliveryServiceFee",
        "ExtraCosts",     # Упаковка + дорога (из формы)
        "PurchasePrice",  # Закуп (из формы)
        "Profit",
        "ProfitPercent"
      ]

      result_rows = [header]

      total_payout    = 0.0
      total_logistics = 0.0
      total_extra     = 0.0
      total_purchase  = 0.0

      sheet.each_with_index do |row, idx|
        next if idx == 0 # header

        data = extract_selected_columns(row)

        barcode = data[:barcode].to_s.strip
        next if barcode.empty?

        reason    = data[:payment_reason].to_s.strip
        payout    = n(data[:payout_amount])
        logistics = n(data[:delivery_service_fee])

        total_payout    += payout
        total_logistics += logistics

        # ✅ КАК У БУХГАЛТЕРА:
        # закуп/прочие расходы проставляем просто по факту "Продажа"
        purchase_value = 0.0
        extra_value    = 0.0

        if reason == SALE_REASON
          unit = @pricing_by_sku[barcode]
          purchase_value = n(unit&.dig(:purchase_price))
          extra_value    = n(unit&.dig(:extra_costs))

          total_purchase += purchase_value
          total_extra    += extra_value
        end

        profit = payout - logistics - extra_value - purchase_value

        denom = extra_value + purchase_value
        profit_percent =
          if denom.positive?
            (profit * 100.0) / denom
          else
            0.0
          end

        result_rows << [
          barcode,
          data[:supplier_article].to_s.strip,
          reason,
          payout.round(2),
          logistics.round(2),
          extra_value.round(2),
          purchase_value.round(2),
          profit.round(2),
          profit_percent.round(2)
        ]
      end

      total_profit = total_payout - total_logistics - total_extra - total_purchase
      total_denom  = total_extra + total_purchase

      total_profit_percent =
        if total_denom.positive?
          (total_profit * 100.0) / total_denom
        else
          0.0
        end

      totals_block = {
        titles: [
          "К перечислению продавцу (итого)",
          "Логистика (итого)",
          "Упаковка + дорога (итого)",
          "Закуп (итого)",
          "Прибыль (как бухгалтер)",
          "Прибыль в процентах % (как бухгалтер)"
        ],
        values: [
          total_payout.round(2),
          total_logistics.round(2),
          total_extra.round(2),
          total_purchase.round(2),
          total_profit.round(2),
          total_profit_percent.round(2)
        ]
      }

      # добавим блок итогов в сам Excel (как в их файле)
      result_rows << []
      result_rows << totals_block[:titles]
      result_rows << totals_block[:values]

      path = export_excel(result_rows)

      { path: path, totals: totals_block }
    end

    private

    def extract_selected_columns(row)
      SELECTED_COLS.transform_values { |idx| row[idx] }
    end

    def n(value)
      return 0.0 if value.nil?
      return value.to_f if value.is_a?(Numeric)

      s = value.to_s.strip
      s = s.gsub(" ", "").tr(",", ".")
      Float(s)
    rescue StandardError
      value.to_f
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
