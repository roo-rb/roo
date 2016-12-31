module Roo
  module Formatters
    module Base
      # converts an integer value to a time string like '02:05:06'
      def integer_to_timestring(content)
        h = (content / 3600.0).floor
        content -= h * 3600
        m = (content / 60.0).floor
        content -= m * 60
        s = content
        Kernel.format("%02d:%02d:%02d", h, m, s)
      end
    end
  end
end
