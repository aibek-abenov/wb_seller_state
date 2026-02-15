class DashboardController < ApplicationController
  before_action :authenticate_user!
  before_action :check_active_subscription!, only: :index

  def index
  end

  # ------------------ FILE UPLOAD ------------------

  def upload
    unless params[:file].present?
      redirect_to dashboard_index_path, alert: "Файл не выбран"
      return
    end

    file = params[:file]

    unless file.content_type.in?(
      ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
    )
      redirect_to dashboard_index_path, alert: "Разрешены только .xlsx файлы"
      return
    end

    # 1️⃣ Сохраняем файл на диск
    stored_path = Uploads::Uploader.call(file: file)

    # 2️⃣ Парсим товары (массив хэшей с ключами sku/name_primary/name_secondary)
    product_names = Excel::Parser.call(file: file)

    # 3️⃣ Большие данные кладём в cache, в session — только ключ
    cache_key = "uploaded_products:#{SecureRandom.hex(12)}"
    Rails.cache.write(cache_key, product_names, expires_in: 30.minutes)

    session[:uploaded_file_path] = stored_path
    session[:uploaded_products_cache_key] = cache_key

    redirect_to pricing_dashboard_index_path
  end

  # ------------------ PRICING FORM ------------------

  def pricing
    cache_key = session[:uploaded_products_cache_key]
    products = cache_key.present? ? Rails.cache.read(cache_key) : nil

    # Вьюха ожидает строковые ключи: product["sku"], product["name_primary"] ...
    @products =
      Array(products).map do |p|
        p.respond_to?(:to_h) ? p.to_h.stringify_keys : p
      end

    if @products.blank?
      redirect_to dashboard_index_path, alert: "Нет загруженного файла"
      return
    end
  end

  # ------------------ SAVE PRICING ------------------

  def save_pricing
    products =
      params.require(:products).values.map do |row|
        {
          sku: row[:sku],
          purchase_price: row[:purchase_price].to_f,
          extra_costs: row[:extra_costs].to_f
        }
      end.reject { |p| p[:sku].blank? }

    file_path = session[:uploaded_file_path]

    unless file_path.present? && ::File.exist?(file_path)
      redirect_to dashboard_index_path, alert: "Файл не найден — загрузите его заново"
      return
    end

    # Processor теперь возвращает { path:, totals: ... }
    result = Excel::Processor.call(
      file_path: file_path,
      pricing: products
    )

    new_file_path = result[:path]
    totals        = result[:totals]

    cleanup_upload!(file_path)

    # перемещаем результат в tmp/exports с безопасным именем
    exports_dir = Rails.root.join("tmp", "exports")
    FileUtils.mkdir_p(exports_dir)

    token = SecureRandom.hex(16)
    download_name = "processed_#{Time.now.strftime("%Y%m%d_%H%M")}.xlsx"
    stored_path = exports_dir.join("#{token}.xlsx").to_s

    FileUtils.mv(new_file_path, stored_path)

    # сохраняем в session данные для скачивания + totals для отображения
    session[:processed_export] = {
      token: token,
      path: stored_path,
      name: download_name,
      totals: totals
    }

    redirect_to export_ready_dashboard_index_path,
                notice: "Файл обработан. Нажмите «Скачать обработанный файл».",
                status: :see_other
  end

  # ------------------ EXPORT READY ------------------

  def export_ready
    data = session[:processed_export]
    path = data && (data["path"] || data[:path])

    unless path.present?
      redirect_to dashboard_index_path, alert: "Нет подготовленного файла для скачивания."
      return
    end

    totals = data && (data["totals"] || data[:totals])

    @totals_titles = totals && (totals["titles"] || totals[:titles]) || []
    @totals_values = totals && (totals["values"] || totals[:values]) || []
  end

  # ------------------ DOWNLOAD ------------------

  def download_processed_export
    data = session[:processed_export]
    path = data && (data["path"] || data[:path])
    name = data && (data["name"] || data[:name]) || "processed.xlsx"

    unless path.present? && ::File.exist?(path)
      head :gone
      return
    end

    file_bytes = ::File.binread(path)

    # одноразово
    session.delete(:processed_export)
    ::File.delete(path) if ::File.exist?(path)

    send_data file_bytes,
              filename: name,
              type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              disposition: "attachment"
  end

  def inactive
  end

  private

  # ------------------ HELPERS ------------------

  def cleanup_upload!(file_path)
    ::File.delete(file_path) if ::File.exist?(file_path)

    if (key = session[:uploaded_products_cache_key]).present?
      Rails.cache.delete(key)
    end

    session.delete(:uploaded_file_path)
    session.delete(:uploaded_products_cache_key)
  end

  def check_active_subscription!
    return if current_user.active?

    redirect_to inactive_dashboard_index_path, alert: "Подписка неактивна"
  end
end
