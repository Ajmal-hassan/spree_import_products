# require_relative 'boot'

require 'rails/all'
require_relative '../../app/models/spree_import_products/configuration.rb'
# require_relative '../../app/models/spree/authentication_method.rb'
# require_relative '../../lib/spree_import_products/engine.rb'
# Require the gems listed in Gemfile, including any gems
# you've limited to :test, :development, or :production.
Bundler.require(*Rails.groups)

module SpreeStarter
  class Application < Rails::Application
    config.autoload_paths += %W(#{config.root}/app/models)
    config.eager_load = true
  end
end
