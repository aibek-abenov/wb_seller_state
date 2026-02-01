Rails.application.routes.draw do
  # ---------- Auth redirects ----------
  authenticated :user do
    root to: "dashboard#index", as: :authenticated_root
  end

  unauthenticated do
    root to: "home#index"
  end

  # ---------- Avo Admin ----------
  mount_avo

  # ---------- Devise ----------
  devise_for :users

  # ---------- Dashboard ----------
  resources :dashboard, only: [:index] do
    collection do
      get :inactive
      post :upload
      get  :pricing
      post :save_pricing

      get  :export_ready
      post :download_processed_export
    end
  end

  # ---------- Healthcheck ----------
  get "up" => "rails/health#show", as: :rails_health_check
end

