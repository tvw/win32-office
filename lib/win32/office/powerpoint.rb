require 'win32/office'
require 'win32/office/chart'


module Win32
  module Office
    module Powerpoint

      module Const
        ppt = nil
        connected = false
        begin
          ppt = WIN32OLE.connect('Powerpoint.Application')
          connected = true
        rescue
          puts "cannot find existing Powerpoint"
          begin
            ppt = WIN32OLE.new("Powerpoint.Application")
          rescue
            puts "cannot start new Powerpoint"
          end
        end

        if ppt
          ppt.DisplayAlerts = false
          WIN32OLE.const_load(ppt, self)
          ppt.Quit unless connected
        end
      end

      class Application < Win32::Office::Wrapper
        def initialize
          logger.debug("Connecting to Powerpoint")
          @obj = WIN32OLE.new('Powerpoint.Application')
          @obj.Visible = true
          @obj.WindowState = 2 # minimize
        end

        def presentations
          Presentations.new(@obj.Presentations)
        end

        def const(name)
          Const.const_get(name)
        end

        def quit
          logger.debug("Quitting Powerpoint")
          presentations.each do |pr|
            pr.Close
          end
          @obj.Quit
          @obj = nil
        end

        def self.quit
          app = Application.new
          app.quit
        end

      end



      class Presentations < Win32::Office::Wrapper
        def initialize(obj)
          @obj = obj
        end

        def add
          logger.debug("Adding a new presentation")
          Presentation.new @obj.Add
        end

        def open(filename)
          logger.debug("Opening presentation #{filename}")
          Presentation.new @obj.Open(File.expand_path(filename), {:WithWindow => Win32::Office::Const::MsoFalse})
        end

        def open_as(filename, asfilename)
          logger.debug("Opening presentation #{filename} as #{asfilename}")
          p = Presentation.new @obj.Open(File.expand_path(filename), {:WithWindow => Win32::Office::Const::MsoFalse})
          p.save_as asfilename
          p
        end
      end



      class Presentation < Win32::Office::Wrapper
        def initialize(obj)
          @obj = obj
        end

        def save_as(filename)
          logger.debug("Saving presentation as #{filename}")
          fullpath = File.expand_path(filename).gsub("/","\\")
          @obj.SaveAs(fullpath)
          self
        end

        def add_slide(slide, no = nil)
          slide.Copy
          @obj.Slides.Paste(no)
        end


        def quit
          logger.debug("Quitting presentation")
          @obj.Close
        end

        def close
          logger.debug("Closing presentation")
          @obj.Save
          @obj.Close
        end
      end

      
      class Chart < Win32::Office::Chart::Chart
        def initialize(shape)
          @shape = shape
          super(@shape.OLEFormat.Object.Application)
        end

        def close
          puts "Closing Chart " + @shape.Name
          super.close
        end
      end


    end
  end
end
