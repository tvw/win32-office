require 'win32/office'

module Win32
  module Office
    module Excel

      module Const
        xl = WIN32OLE.new('Excel.Application')
        xl.DisplayAlerts = false
        xl.Visible = false
        WIN32OLE.const_load(xl, self)
        xl.ole_free
      end

      class Application < Win32::Office::Wrapper
        def initialize

          @obj = nil
          @connect = false
          begin
            @obj = WIN32OLE.connect('Excel.Application')
            @connect = true
          rescue
            puts "cannot find existing Excel"
            begin
              @obj = WIN32OLE.new("Excel.Application")
            rescue
              puts "cannot start new Excel"
            end
          end

          unless @obj
            raise "Excel not found"
          end
        end

        # Returns true, if we connected to a running Excel
        # instead of creating a new one.
        def connect?
          @connect
        end

        def workbooks
          Workbooks.new(@obj.Workbooks)
        end

        def const(name)
          Const.const_get(name)
        end

        # Quit Excel, unless we connected to a running instance.
        def quit
          @obj.Quit unless connect?
        end

      end

      class Workbooks < Win32::Office::Wrapper
        def initialize(obj)
          @obj = obj
        end

        def add
          Workbook.new @obj.Add
        end

        def open(filename)
          Workbook.new @obj.Open(File.expand_path(filename).gsub("/",'\\'))
        end

        def open_as(filename, asfilename)
          p = Workbook.new @obj.Open(File.expand_path(filename))
          p.save_as asfilename
          p
        end
      end

      class Workbook < Win32::Office::Wrapper
        def initialize(obj)
          @obj = obj
        end

        def save_as(filename)
          @obj.SaveAs(File.expand_path(filename))
          self
        end

#        def add_slide(slide, no = nil)
#          slide.Copy
#          @obj.Slides.Paste(no)
#        end


        def quit
          @obj.Close
        end

        def close
          @obj.Save
          @obj.Close
        end
      end




    end
  end
end
