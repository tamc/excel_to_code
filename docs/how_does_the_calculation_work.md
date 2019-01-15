In the generated C file, calculations are done when you get a value, rather than when you set a value.

The results are then cached, so you need to call reset() when you later set a value.

i.e.,

set_control_a1(new_value) # No calculations done
set_control_a2(new_value) # No calculations done
set_control_a3(new_value) # No calculations done

output_a1() # Calculations done in all cells that a1 depends, results cached
output_a2() # Calculations done in all cells that a2 depends, that haven’t already been calculated for output_a1

set_control_a1(new_value) # DONT DO THIS! The changed value won’t be picked up

reset() # All cached calculations erased

set_control_a1(new_value) # Now you can do this

output_a1() # Calculations done, based on new value of a1
