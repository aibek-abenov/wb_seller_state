class Avo::Resources::User < Avo::BaseResource
  self.title = :email
  self.includes = []
  # self.attachments = []
  # self.search = {
  #   query: -> { query.ransack(id_eq: q, m: "or").result(distinct: false) }
  # }

  def fields
    field :email, as: :text
    field :role, as: :text
    field :active_until, as: :date
  end
end
