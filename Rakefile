begin
  require 'bones'
rescue LoadError
  puts '### Please install the "bones" gem ###'
end

ensure_in_path 'lib'
require 'roo'

task :default => 'test:run'
# task 'gem:release' => 'test:run'

Bones {
  name  'roo'
  authors  'Thomas Preymesser'
  email  'thopre@gmail.com'
  url  'http://roo.rubyforge.org/'
  version  Roo::VERSION
  depend_on 'spreadsheet', '> 0.6.4'
  #--
  # rel. 0.6.4 causes an invalid Date error if we
  # have a datetime value of 2006-02-02 10:00:00
  #++
  depend_on 'nokogiri' #, '>= 0.0.1'
  #TODO: brauchen wir das noch? depend_on 'gimite-google-spreadsheet-ruby','>= 0.0.5'
  #depend_on 'febeling-rubyzip','>= 0.9.2' # meine aktuelle Version
  #TODO: warum brauchen wir das? es lief doch auch vorher ohne dieses spezielle gem
  depend_on 'rubyzip' # rubyzip wird benoetigt
  depend_on 'google-spreadsheet-ruby'
  depend_on 'choice'
  depend_on 'todonotes'
}

# EOF
