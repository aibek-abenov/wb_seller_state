require "roo"
require "caxlsx"
require "bigdecimal"

module Excel
  class Processor
    SALE_REASON = "Продажа".freeze
    RETURN_REASON = "Возврат".freeze
    LOGISTICS_REASON = "Логистика".freeze

    VALID_PAYMENT_REASONS = [
      "Возврат",
      "Логистика",
      "Продажа",
      "Удержание",
      "Хранение",
      "Штраф"
    ].freeze

    # Roo индексы — с нуля
    BARCODE_INDEX = 8   # I  — Баркод
    DOCUMENT_TYPE_INDEX = 9 # J - Тип документа
    REASON_INDEX  = 10  # K  — Обоснование для оплаты
    DELIVERY_FEE_INDEX  = 36 # AK — Услуги по доставке товара покупателю (логистика)

    SELECTED_COLS = {
      supplier_article:      5,   # F  — Артикул поставщика
      barcode:               BARCODE_INDEX,
      payment_reason:        REASON_INDEX,
      payout_amount:         33,  # AH — К перечислению продавцу...
      delivery_service_fee:  36   # AK — Логистика
    }.freeze

    HEADER = [
      "Баркод",
      "Артикул поставщика",
      "Обоснование для оплаты",
      "К перечислению продавцу",
      "Услуги по доставке (логистика)",
      "Упаковка или прочие расходы",
      "Закуп",
      "Прибыль",
      "Прибыль в процентах %"
    ].freeze

    TOTALS_TITLES = [
      "К перечислению продавцу (итого)",
      "Логистика (итого)",
      "Упаковка или прочие расходы",
      "Закуп (итого)",
      "Прибыль (как бухгалтер)",
      "Прибыль в процентах %"
    ].freeze

    def self.call(file_path:, pricing:)
      new(file_path, pricing).call
    end

    def initialize(file_path, pricing)
      @file_path = file_path
      @pricing_by_sku = pricing.index_by { |p| p[:sku].to_s }
    end

    # @return { path: "/tmp/xxx.xlsx", totals: { titles: [...], values: [...] } }
    def call
      filtered_rows = stream_and_filter
      result_rows, totals = calculate(filtered_rows)
      path = export_excel(result_rows)

      { path: path, totals: totals }
    end

    private

    # Однопроходная обработка: фильтрация + merge логистики
    def stream_and_filter
      workbook = Roo::Excelx.new(@file_path)
      sheet = workbook.sheet(0)

      pending_logistics = {}
      filtered_rows = []

      first = true
      sheet.each_row_streaming do |row|
        if first
          first = false
          next
        end

        barcode_val = cell_value(row, BARCODE_INDEX).to_s.strip
        next if barcode_val.empty?

        reason_val = cell_value(row, REASON_INDEX).to_s.strip
        next unless VALID_PAYMENT_REASONS.include?(reason_val)

        doc_type_val = cell_value(row, DOCUMENT_TYPE_INDEX).to_s.strip

        row_data = {
          barcode: barcode_val,
          supplier_article: cell_value(row, SELECTED_COLS[:supplier_article]).to_s.strip,
          payment_reason: reason_val,
          payout_amount: cell_value(row, SELECTED_COLS[:payout_amount]),
          delivery_service_fee: cell_value(row, SELECTED_COLS[:delivery_service_fee]),
          doc_type: doc_type_val
        }

        is_logistics = reason_val == LOGISTICS_REASON && doc_type_val.empty?
        is_sale = reason_val == SALE_REASON && doc_type_val == SALE_REASON

        if is_logistics
          pending_logistics[barcode_val] =
            n(pending_logistics[barcode_val]) + n(row_data[:delivery_service_fee])
          next
        end

        if is_sale && pending_logistics.key?(barcode_val)
          row_data[:delivery_service_fee] =
            n(row_data[:delivery_service_fee]) + pending_logistics.delete(barcode_val)
        end

        filtered_rows << row_data
      end

      # Оставшиеся логистики без продажи — как отдельные строки
      pending_logistics.each do |barcode, fee|
        filtered_rows << {
          barcode: barcode,
          supplier_article: "",
          payment_reason: LOGISTICS_REASON,
          payout_amount: BigDecimal("0"),
          delivery_service_fee: fee,
          doc_type: ""
        }
      end

      filtered_rows
    ensure
      workbook&.close if workbook.respond_to?(:close)
    end

    # Расчёт прибыли и формирование result_rows + totals
    def calculate(filtered_rows)
      result_rows = [HEADER]

      total_payout    = BigDecimal("0")
      total_logistics = BigDecimal("0")
      total_extra     = BigDecimal("0")
      total_purchase  = BigDecimal("0")

      filtered_rows.each do |data|
        barcode = data[:barcode]
        reason  = data[:payment_reason]

        payout = n(data[:payout_amount])
        payout = -payout if reason == RETURN_REASON && payout.positive?

        logistics = n(data[:delivery_service_fee])

        total_payout    += payout
        total_logistics += logistics

        purchase_value = BigDecimal("0")
        extra_value    = BigDecimal("0")

        if reason == SALE_REASON
          unit = @pricing_by_sku[barcode]
          purchase_value = n(unit&.dig(:purchase_price))
          extra_value    = n(unit&.dig(:extra_costs))

          total_purchase += purchase_value
          total_extra    += extra_value
        end

        profit = payout - logistics - extra_value - purchase_value
        profit_percent = percent(profit, extra_value + purchase_value)

        result_rows << [
          barcode,
          data[:supplier_article],
          reason,
          payout.to_f.round(2),
          logistics.to_f.round(2),
          extra_value.to_f.round(2),
          purchase_value.to_f.round(2),
          profit.to_f.round(2),
          profit_percent.to_f.round(2)
        ]
      end

      total_profit = total_payout - total_logistics - total_extra - total_purchase
      total_profit_percent = percent(total_profit, total_extra + total_purchase)

      totals_block = {
        titles: TOTALS_TITLES,
        values: [
          total_payout.to_f.round(2),
          total_logistics.to_f.round(2),
          total_extra.to_f.round(2),
          total_purchase.to_f.round(2),
          total_profit.to_f.round(2),
          total_profit_percent.to_f.round(2)
        ]
      }

      result_rows << []
      result_rows << totals_block[:titles]
      result_rows << totals_block[:values]

      [result_rows, totals_block]
    end

    def cell_value(row, index)
      cell = row[index]
      cell.respond_to?(:value) ? cell.value : cell
    end

    # Безопасно приводим к BigDecimal
    def n(value)
      return BigDecimal("0") if value.nil?
      return BigDecimal(value, Float::DIG) if value.is_a?(Numeric)

      s = value.to_s.strip.gsub(" ", "").tr(",", ".")
      BigDecimal(s)
    rescue ArgumentError, TypeError
      BigDecimal("0")
    end

    def percent(profit, denom)
      denom.positive? ? (profit * 100) / denom : BigDecimal("0")
    end

    def export_excel(rows)
      exports_dir = Rails.root.join("tmp", "exports")
      FileUtils.mkdir_p(exports_dir)
      path = exports_dir.join("result_#{SecureRandom.hex(8)}.xlsx").to_s
      tmp_path = "#{path}.tmp"

      Axlsx::Package.new(use_shared_strings: false) do |p|
        p.workbook.add_worksheet(name: "Result") do |sheet|
          rows.each { |r| sheet.add_row(r) }
        end
        p.serialize(tmp_path)
      end

      FileUtils.mv(tmp_path, path)
      path
    end
  end
end
