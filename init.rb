# Include hook code here
require "railsworksheet"

ActionView::Template.register_template_handler 'rxls', Worksheet::Handler