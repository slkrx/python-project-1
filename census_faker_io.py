def get_input(function, prompt, args=None):
    while True:
        try:
            return function(input(prompt), args) if args else function(input(prompt))
        except ValueError as e:
            print(e)

def dropdown_validate(s, allowable_values):
    if s in allowable_values:
        return s
    else:
        raise ValueError("Must be one of the following values: " + ", ".join(allowable_values))

def s_to_i(s):
    try:
        n = int(s)
    except ValueError:
        raise ValueError("Not an integer")
    if n < 1:
        raise ValueError("Value must be a positive integer")
    else:
        return n

def s_to_b(s):
    if s == "y":
        return True
    elif s == "n":
        return False
    else:
        raise ValueError("Must enter \"y\" or \"n\"")

def get_boolean(prompt):
    return get_input(s_to_b, prompt)

def get_positive_integer(prompt):
    return get_input(s_to_i, prompt)

def get_dropdown_input(prompt, allowable_values):
    return get_input(dropdown_validate, prompt, allowable_values)