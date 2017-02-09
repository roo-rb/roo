module Roo
  class Excelx
    class Cell
      class Base
        attr_reader :cell_type, :cell_value, :value

        # FIXME: I think style should be deprecated. Having a style attribute
        #        for a cell doesn't really accomplish much. It seems to be used
        #        when you want to export to excelx.
        attr_reader :style


        # FIXME: Updating a cell's value should be able tochange the cell's type,
        #        but that isn't currently possible. This will cause weird bugs
        #        when one changes the value of a Number cell to a String. e.g.
        #
        #           cell = Cell::Number(*args)
        #           cell.value = 'Hello'
        #           cell.formatted_value # => Some unexpected value
        #
        #        Here are two possible solutions to such issues:
        #        1. Don't allow a cell's value to be updated. Use a method like
        #          `Sheet.update_cell` instead. The simple solution.
        #        2. When `cell.value = ` is called, use injection to try and
        #           change the type of cell on the fly. But deciding what type
        #           of value to pass to `cell.value=`. isn't always obvious. e.g.
        #           `cell.value = Time.now` should convert a cell to a DateTime,
        #           not a Time cell. Time cells would be hard to recognize because
        #           they are integers. This approach would require a significant
        #           change to the code as written. The complex solution.
        #
        #        If the first solution is used, then this method should be
        #        deprecated.
        attr_writer :value

        def initialize(value, formula, excelx_type, style, link, coordinate)
          @link = !!link
          @cell_value = value
          @cell_type = excelx_type
          @formula = formula
          @style = style
          @coordinate = coordinate
          @type = :base
          @value = link? ? Roo::Link.new(link, value) : value
        end

        def type
          if formula?
            :formula
          elsif link?
            :link
          else
            @type
          end
        end

        def formula?
          !!@formula
        end

        def link?
          !!@link
        end

        alias_method :formatted_value, :value

        def to_s
          formatted_value
        end

        # DEPRECATED: Please use link instead.
        def hyperlink
          warn '[DEPRECATION] `hyperlink` is deprecated.  Please use `link` instead.'
        end

        # DEPRECATED: Please use cell_value instead.
        def excelx_value
          warn '[DEPRECATION] `excelx_value` is deprecated.  Please use `cell_value` instead.'
          cell_value
        end

        # DEPRECATED: Please use cell_type instead.
        def excelx_type
          warn '[DEPRECATION] `excelx_type` is deprecated.  Please use `cell_type` instead.'
          cell_type
        end

        def empty?
          false
        end
      end
    end
  end
end
