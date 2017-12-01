def g(x, z):
    print x,z
    x.append(z)
    print x
    return x

y = [1, 2, 3]
g(y, 4)
y.extend(g(y[:], 4))
print y
y = [1, 2, 3]
a = g(y[:], 4)
y.extend(g(y, 4))
print y