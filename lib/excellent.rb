require 'excellent/base_overrides'
require 'excellent/template_handler'

Mime::Type.register "application/vnd.ms-excel", :xls

ActionView::Template.register_template_handler :xerb, ActionView::TemplateHandlers::XERB

