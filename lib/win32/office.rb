require 'win32ole'
require 'logger'
require 'office/version'

module Win32
  module Office
    module Const
      xl = WIN32OLE.new('Excel.Application')
      xl.DisplayAlerts = false
      xl.visible = false
      WIN32OLE.const_load("Microsoft Office #{xl.Version} Object Library", self)
      xl.ole_free
    end

    LOGGER = Logger.new(STDERR)

    class Wrapper
      def logger
        LOGGER
      end

      def method_missing(sym, *args, &block)
#        puts "Sending #{sym}(#{args.join(',')}) to obj"
        @obj.__send__(sym, *args, &block)
      end
    end

    module Base
      def logger
        LOGGER
      end

      def method_missing(sym, *args, &block)
        @obj.__send__(sym, *args, &block)
      end

      def obj
        @obj
      end
    end


  end
end
