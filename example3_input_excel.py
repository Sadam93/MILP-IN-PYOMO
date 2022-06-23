import matplotlib.pyplot as plt
import pandas as pd
import pyomo.environ as py
from pyomo.environ import *

# model
# from pyomo.core import value ,

model = py.ConcreteModel()
data = pd.read_excel('input_data.xlsx', sheet_name='set')

# Set
model.t = py.Set(initialize=data.time)
t = model.t

# Parameter
model.pb = py.Param(t, initialize=dict(zip(data.time.values, data.base_load1.values)), doc='Base load')
model.c = py.Param(t, initialize=dict(zip(data.time.values, data.price.values)), doc='cost')
model.dt = py.Param(initialize=1, doc='time duration ')
dt = model.dt
pb = model.pb
c = model.c
# Variable
model.pl = py.Var(t, bounds=(0, 6))

pl = model.pl

# Constraints
model.c1 = py.Constraint(expr=sum([pl[i] for i in t]) * dt == 8)

# Objective
sum_ob = sum((pb[i] + pl[i]) * c[i] for i in t)
model.obj = py.Objective(expr=sum_ob * dt)

# Solve
opt = py.SolverFactory('cplex')
result = opt.solve(model)
# print(result)
# model.pprint()
model.display()
# Print variables
print('------------Countable variable Solution------------------------')
for i in model.t:
    print('x[%i] = %i' % (i, py.value(pl[i])))
print('objective function = ', py.value(model.obj))

# for ll in model.t:
#   pl_xls = (model.pl.pprint())
# pl_xls1 = pd.DataFrame(pl_xls)
# pl_xls1.to_excel('Results.xlsx', sheet_name = 'time')
# G_Pur = (i, py.value(model.Gp[i]))
# G_Pur = pd.DataFrame(G_Pur)
# G_Pur.to_excel('Results.xlsx', sheet_name = 'Gp')
# with pd.ExcelWriter('Results.xlsx') as writer:
#   E_Battery.to_excel(writer, sheet_name='Eb')
#   G_Pur.to_excel(writer, sheet_name='Gp')


# Plotting

# Plotting
time = []
time_head = []
pbase = []
pload = []
pload_head = []
cost = []

for ll in range(len([t, pb, pl, c])):
    if ll == 0:
        for i in t:
            time.append(value(t[i]))
    elif ll == 1:
        for i in pb:
            pbase.append(value(pb[i]))
    elif ll == 2:
        for i in pl:
            pload.append(value(pl[i]))
    else:
        for i in c:
            cost.append(value(c[i]))

fig, (ax0, ax1) = plt.subplots(nrows=2, ncols=1)

ax0.plot(time, pbase, 'b-', label='Base')
ax0.plot(time, pload, 'g--', label='Controable Load')
ptotal = [pbase[i] + pload[i] for i in range(len(pbase))]
ax0.plot(time, ptotal, 'r', label='Total Load')
ax0.set_ylabel('Electric power (W)')
ax0.legend(loc='upper left')
ax1.plot(time, cost, 'y-+', label='Cost')
ax1.legend(loc='upper left')
ax1.set_xlabel('Time (hour)')
ax1.set_ylabel('Electricity price ($/W/h)')
plt.show()

time_excel = pd.DataFrame({'time': time})
writer = pd.ExcelWriter('Results.xlsx', engine='xlsxwriter')
time_excel.to_excel(writer, sheet_name='Data', startcol=0, index=False, header=True)
pload_excel = pd.DataFrame({'controllable load': pload})
pload_excel.to_excel(writer, sheet_name='Data', startcol=1, index=False, header=True)
writer.save()
