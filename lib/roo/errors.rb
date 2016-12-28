module Roo
  # A base error class for Roo. Most errors thrown by Roo should inherit from
  # this class.
  class Error < StandardError; end

  # Raised when Roo cannot find a header row that matches the given column
  # name(s).
  class HeaderRowNotFoundError < Error; end

  class FileNotFound < Error; end
end
