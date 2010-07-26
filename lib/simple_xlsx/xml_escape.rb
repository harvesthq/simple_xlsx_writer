unless String.method_defined? :to_xs
  require 'fast_xs' #dep
  class String
    alias_method :to_xs, :fast_xs
  end
end
