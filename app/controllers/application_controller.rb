class ApplicationController < ActionController::Base
  protected

  def after_sign_out_path_for(_resource_or_scope)
    root_path
  end

  def after_sign_in_path_for(_resource)
    dashboard_index_path
  end
end
