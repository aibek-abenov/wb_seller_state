class Rack::Attack
  # Limit file uploads to 5 per minute per IP
  throttle("uploads/ip", limit: 5, period: 60) do |req|
    req.ip if req.path == "/dashboard/upload" && req.post?
  end

  # Limit pricing submissions to 5 per minute per IP
  throttle("save_pricing/ip", limit: 5, period: 60) do |req|
    req.ip if req.path == "/dashboard/save_pricing" && req.post?
  end

  # General rate limit: 60 requests per minute per IP (except assets)
  throttle("requests/ip", limit: 60, period: 60) do |req|
    req.ip unless req.path.start_with?("/assets")
  end

  self.throttled_responder = lambda do |req|
    [429, { "Content-Type" => "text/plain" }, ["Too many requests. Please try again later.\n"]]
  end
end
