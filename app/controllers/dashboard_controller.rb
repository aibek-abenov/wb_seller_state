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

    if file.size > 10.megabytes
      redirect_to dashboard_index_path, alert: "Файл слишком большой (макс. 10 МБ)"
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
    # Не позволяем ставить в очередь, если уже есть активная задача
    if session[:processing_job_token].present?
      existing = Rails.cache.read("excel_job:#{session[:processing_job_token]}")
      if existing && existing[:status] == "processing"
        redirect_to processing_dashboard_index_path, alert: "У вас уже есть файл в обработке"
        return
      end
    end

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

    job_token = SecureRandom.hex(16)

    ExcelProcessingJob.perform_later(
      user_id: current_user.id,
      file_path: file_path,
      pricing: products,
      job_token: job_token
    )

    # Очищаем upload-данные из session (файл удалит job)
    if (key = session[:uploaded_products_cache_key]).present?
      Rails.cache.delete(key)
    end
    session.delete(:uploaded_file_path)
    session.delete(:uploaded_products_cache_key)

    session[:processing_job_token] = job_token
    redirect_to processing_dashboard_index_path
  end

  # ------------------ PROCESSING STATUS ------------------

  def processing
    @job_token = session[:processing_job_token]
    unless @job_token.present?
      redirect_to dashboard_index_path, alert: "Нет активных задач"
      return
    end
  end

  def job_status
    job_token = session[:processing_job_token]
    raw = Rails.cache.read("excel_job:#{job_token}")
    status_data = raw.is_a?(Hash) ? raw.symbolize_keys : { status: "unknown" }

    if status_data[:status] == "completed"
      totals = status_data[:totals]
      totals = totals.symbolize_keys if totals.is_a?(Hash)

      session[:processed_export] = {
        token: job_token,
        path: status_data[:path],
        name: status_data[:name],
        totals: totals
      }
      session.delete(:processing_job_token)
      Rails.cache.delete("excel_job:#{job_token}")
    end

    render json: { status: status_data[:status], error: status_data[:error] }
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

  def check_active_subscription!
    return if current_user.active?

    redirect_to inactive_dashboard_index_path, alert: "Подписка неактивна"
  end
end
