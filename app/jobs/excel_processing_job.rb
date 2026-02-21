class ExcelProcessingJob < ApplicationJob
  queue_as :default

  retry_on StandardError, wait: 5.seconds, attempts: 2
  discard_on ActiveJob::DeserializationError

  def perform(user_id:, file_path:, pricing:, job_token:)
    cache_key = "excel_job:#{job_token}"

    Rails.cache.write(cache_key, { status: "processing" }, expires_in: 1.hour)

    result = Excel::Processor.call(
      file_path: file_path,
      pricing: pricing
    )

    new_file_path = result[:path]
    totals        = result[:totals]

    exports_dir = Rails.root.join("tmp", "exports")
    FileUtils.mkdir_p(exports_dir)

    stored_path = exports_dir.join("#{job_token}.xlsx").to_s
    FileUtils.mv(new_file_path, stored_path)

    ::File.delete(file_path) if ::File.exist?(file_path)

    Rails.cache.write(cache_key, {
      status: "completed",
      path: stored_path,
      name: "processed_#{Time.now.strftime('%Y%m%d_%H%M')}.xlsx",
      totals: totals
    }, expires_in: 1.hour)

    Turbo::StreamsChannel.broadcast_replace_to(
      "excel_processing_#{user_id}",
      target: "processing_status",
      partial: "dashboard/processing_complete"
    )
  rescue => e
    Rails.cache.write(cache_key, {
      status: "failed",
      error: e.message
    }, expires_in: 1.hour)

    ::File.delete(file_path) if file_path && ::File.exist?(file_path)

    Turbo::StreamsChannel.broadcast_replace_to(
      "excel_processing_#{user_id}",
      target: "processing_status",
      partial: "dashboard/processing_failed",
      locals: { error: e.message }
    )

    raise
  end
end
