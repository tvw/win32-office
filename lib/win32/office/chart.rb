require 'win32/office'

module Win32
  module Office

    module Chart

      class Application
        include Win32::Office::Base

        def initialize(obj)
          @obj = obj
        end

        def close
          obj.Update
          obj.Quit
        end

        def app
          self
        end
      end


      class Chart
        include Win32::Office::Base

        def initialize(app)
          @app = Application.new(app)
          @obj = @app.Chart
          @datasheet = Datasheet.new(@app)
          if block_given?
            yield self
            self.close
          end
        end

        def datasheet
          yield @datasheet if block_given?
          @datasheet
        end

        def close
          @app.close
        end
      end


      class Datasheet
        include Win32::Office::Base

        def initialize(app)
          @obj = app.DataSheet
        end

        def clear(keeprows = nil, keepcols = nil, maxrows=50, maxcols=50)
          cells = obj.Cells

          if keeprows or keepcols
            if keeprows
              keepcols = 0 unless keepcols
              ((keeprows+1)..maxrows).each do |r|
                ((keepcols+1)..maxcols).each do |c|
                  cells(r,c).ClearContents
                end
              end
            end
          else
            cells.ClearContents
          end

          self
        end

      end

    end



  end
end
