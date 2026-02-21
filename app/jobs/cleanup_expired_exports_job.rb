class CleanupExpiredExportsJob < ApplicationJob
  queue_as :default

  def perform
    cleanup_directory(Rails.root.join("tmp", "exports"), max_age: 2.hours)
    cleanup_directory(Rails.root.join("storage", "uploads"), max_age: 1.hour)
  end

  private

  def cleanup_directory(dir, max_age:)
    return unless dir.exist?

    Dir.glob(dir.join("*")).each do |file|
      ::File.delete(file) if ::File.mtime(file) < max_age.ago
    rescue Errno::ENOENT
      # File already deleted by another process
    end
  end
end
