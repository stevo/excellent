#require 'erb'
require 'excellent/xhtml2xls'

module ActionView
  module TemplateHandlers
    class XERB  < ERB
      include Compilable

      cattr_accessor :erb_trim_mode
      self.erb_trim_mode = '-'

      def compile(template)
        src = %{
        ts = %{#{template.source}}
        xhtml = self.render(:inline => ts, :type => :erb )
        h2x = Hsh2Xls.new(xhtml)
        h2x.output
        }

        RUBY_VERSION >= '1.9' ? src.sub(/\A#coding:.*\n/, '') : src
      end
    end
  end
end