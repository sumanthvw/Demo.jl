module Demo

export ffill, greet

greet() = print("Hello World!")

ffill(v) = v[accumulate(max, [i*!ismissing(v[i]) for i in 1:length(v)], init=1)]


end # module Demo
