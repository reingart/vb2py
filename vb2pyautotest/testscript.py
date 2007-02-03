from Globals import Factorial
results = []
for x in (0, 1, 2, 5, 10):
	results.append((Factorial(x), x))
f = open('test_Globals_Factorial_py.txt', 'w')
f.write('# vb2Py Autotest results\n')
f.write('\n'.join([', '.join(map(str, x)) for x in results]))
f.close()
from Globals import Add
results = []
for x in (-10.0, -5.5, -1.5, 0.0, 1.5, 5.5, 10.0):
	for y in (-10.0, -5.5, -1.5, 0.0, 1.5, 5.5, 10.0):
		results.append((Add(x,y), x,y))
f = open('test_Globals_Add_py.txt', 'w')
f.write('# vb2Py Autotest results\n')
f.write('\n'.join([', '.join(map(str, x)) for x in results]))
f.close()