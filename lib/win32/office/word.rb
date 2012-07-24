require 'win32/office'

module Win32
  module Office
    module Word

      module Const
        wd = WIN32OLE.new('Word.Application')
        wd.DisplayAlerts = false
        wd.Visible = false
        WIN32OLE.const_load(wd, self)
        wd.Quit
      end

    end
  end
end
