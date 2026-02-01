class User < ApplicationRecord
  # Include default devise modules. Others available are:
  # :confirmable, :lockable, :timeoutable, :trackable and :omniauthable
  devise :database_authenticatable, :registerable,
    :recoverable, :rememberable, :validatable, :timeoutable

  # enum role: { user: "user", admin: "admin" }

  def active?
    return true if admin?

    active_until.present? && active_until >= Date.today
  end

  def admin?
    role == "admin"
  end
end
