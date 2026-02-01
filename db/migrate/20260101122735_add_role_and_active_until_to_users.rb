class AddRoleAndActiveUntilToUsers < ActiveRecord::Migration[8.0]
  def change
    add_column :users, :role, :string, null: false, default: "user"
    add_column :users, :active_until, :date
  end
end
