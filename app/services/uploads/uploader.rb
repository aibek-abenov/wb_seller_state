module Uploads
  class Uploader
    UPLOAD_DIR = Rails.root.join("storage", "uploads")

    def self.call(file:)
      FileUtils.mkdir_p(UPLOAD_DIR)

      stored_path =
        UPLOAD_DIR.join("#{SecureRandom.uuid}_#{file.original_filename}")

      FileUtils.cp(file.path, stored_path)

      stored_path.to_s
    end
  end
end
