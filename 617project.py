import pandas as pd
import gurobipy as gp
from gurobipy import GRB
from itertools import product


Factory = ['Troy, NY', 'Newark, NJ', 'Harrisburg, PA']
Line = ['line 1', 'line 2', 'line 3']
Month = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


# # Model 1

m = gp.Model("All same")


# Transportation cost
ship = pd.read_excel('Project Data.xlsx', 'shipping_cost')
ship = ship.iloc[:, 1:]
c = {(facility, customer): ship.iloc[facility, customer]
            for facility in range(0, 3)
            for customer in range(0, 8)}

# Production rate of production line l at factory i (Units per hour) 
wear = pd.read_excel('Project Data.xlsx', 'production_rate')
wear = wear.iloc[:, 1:]
w = {(facility, line): wear.iloc[facility, line]
            for facility in range(0, 3)
            for line in range(0, 3)}

# Production capacity : M

M = {(facility, line): wear.iloc[facility, line] * 160
            for facility in range(0, 3)
            for line in range(0, 3)}
# energy cost
energy = pd.read_excel('Project Data.xlsx', 'energy_cost')
energy = energy.iloc[:, 1:]
e = {(facility, line): energy.iloc[facility, line]
            for facility in range(0, 3)
            for line in range(0, 3)}
# demand
demand = pd.read_excel('Project Data.xlsx', 'demand')
demand = demand.iloc[:, 1:]
d = {(month, customer): demand.iloc[month, customer]
            for month in range(0, 12)
            for customer in range(0, 8)}


# ### Decision variable



# the number of units transported from factory i to store j in month m
u = m.addVars(list(product(range(0, 3), range(0, 8), range(0, 12))),lb = 0 ,vtype = GRB.INTEGER, name='units transported')

# the number of units produced at factory i from production line l in month m
s = m.addVars(list(product(range(0, 3), range(0, 3), range(0, 12))),lb = 0,vtype = GRB.INTEGER, name='units produced')


# ### Objective function
m.setObjective(( gp.quicksum(1.2 *c[i,j] * u[i, j, m] + e[i, l] * s[i, l, m] for l in range(0, 3) 
                             for m in range(0, 12) for j in range(0, 8) for i in range(0, 3))), GRB.MINIMIZE)


# ### Constractions
# 1. Demand Fulfilment on Transportation
m.addConstrs((gp.quicksum(u[i, j, m] for i in range(0, 3)) >= d[m, j] 
              for j in range(0, 8) for m in range(0, 12)), name='Demand')

# 2. Production Capacity
m.addConstrs((s[i, l, m] <= M[i, l] 
              for l in range(0, 3) for i in range(0, 3) for m in range(0, 12)), name='Capacity')

# 3. Demand Fulfilment on Factory
m.addConstrs((gp.quicksum(s[i, l, m] for i in range(0, 3) for l in range(0, 3)) >= (gp.quicksum(d[m, j] for j in range(0, 8)))
              for m in range(0, 12)), name='Demand')

# 4. Transportation units can’t exceed production units
m.addConstrs((gp.quicksum(u[i, j, m] for j in range(0, 8)) <= gp.quicksum(s[i, l, m] for l in range(0, 3))
              for i in range(0, 3) for m in range(0, 12)), name='Transport units')

# 5. Consistent production
m.addConstrs((gp.quicksum(s[0, l, m] for l in range(0, 3)) == gp.quicksum(s[1, l, m] for l in range(0, 3))
              for m in range(0, 12)), name = 'Consist production 1')


m.addConstrs((gp.quicksum(s[1, l, m] for l in range(0, 3)) == gp.quicksum(s[2, l, m] for l in range(0, 3))
              for m in range(0, 12)), name = 'Consist production 2')

# 6. Equal wear and tear
m.addConstrs((s[i, 0, m] / w[i, 0] == s[i, 1, m] / w[i, 1] 
              for i in range(0, 3) for m in range(0, 12)), name = '1')
m.addConstrs((s[i, 1, m] / w[i, 1] == s[i, 2, m] / w[i, 2] 
              for i in range(0, 3) for m in range(0, 12)), name = '2')


# Optimization
m.optimize()


obj = m.getObjective()
model1_price = obj.getValue()
model1_price

# Output
factory = []
production_line = []
month = []
quantity = []
for facility in s.keys():
    factory.append(Factory[facility[0]])
    production_line.append(Line[facility[1]])
    month.append(Month[facility[2]])
    quantity.append(s[facility].x)

model1 = pd.DataFrame({'Factory':factory, 'ProductionLine':production_line,'Month':month, 'Quantity':quantity})
model1.to_excel('Model1_same_quantity.xlsx', index = 0)


# # Model 2

# In[11]:


m = gp.Model("Difference")




# Transportation cost
# shipping_cost
ship = pd.read_excel('Project Data.xlsx', 'shipping_cost')
ship = ship.iloc[:, 1:]
c = {(facility, customer): ship.iloc[facility, customer]
            for facility in range(0, 3)
            for customer in range(0, 8)}
# Production rate of production line l at factory i (Units per hour) 
wear = pd.read_excel('Project Data.xlsx', 'production_rate')
wear = wear.iloc[:, 1:]
w = {(facility, line): wear.iloc[facility, line]
            for facility in range(0, 3)
            for line in range(0, 3)}
# Production capacity : M

M = {(facility, line): wear.iloc[facility, line] * 160
            for facility in range(0, 3)
            for line in range(0, 3)}
# energy cost
energy = pd.read_excel('Project Data.xlsx', 'energy_cost')
energy = energy.iloc[:, 1:]
e = {(facility, line): energy.iloc[facility, line]
            for facility in range(0, 3)
            for line in range(0, 3)}
# demand
demand = pd.read_excel('Project Data.xlsx', 'demand')
demand = demand.iloc[:, 1:]
d = {(month, customer): demand.iloc[month, customer]
            for month in range(0, 12)
            for customer in range(0, 8)}


# ### Decision variable

# the number of units transported from factory i to store j in month m
u = m.addVars(list(product(range(0, 3), range(0, 8), range(0, 12))),lb = 0 ,vtype = GRB.INTEGER, name='units transported')

# the number of units produced at factory i from production line l in month m
s = m.addVars(list(product(range(0, 3), range(0, 3), range(0, 12))),lb = 0,vtype = GRB.INTEGER, name='units produced')


# ### Objective function



m.setObjective((gp.quicksum(1.2 * c[i,j] * u[i, j, m] + e[i, l] * s[i, l, m] for l in range(0, 3) 
                             for m in range(0, 12) for j in range(0, 8) for i in range(0, 3))), GRB.MINIMIZE)


# ### Constractions



# 1. Demand Fulfilment on Transportation
m.addConstrs((gp.quicksum(u[i, j, m] for i in range(0, 3)) >= d[m, j] 
              for j in range(0, 8) for m in range(0, 12)), name='Demand')

# 2. Production Capacity
m.addConstrs((s[i, l, m] <= M[i, l] 
              for l in range(0, 3) for i in range(0, 3) for m in range(0, 12)), name='Capacity')

# 3. Demand Fulfilment on Factory
m.addConstrs((gp.quicksum(s[i, l, m] for i in range(0, 3) for l in range(0, 3)) >= (gp.quicksum(d[m, j] 
            for j in range(0, 8))) for m in range(0, 12)), name='Demand')

# 4. Transportation units can’t exceed production units
m.addConstrs((gp.quicksum(u[i, j, m] for j in range(0, 8)) <= gp.quicksum(s[i, l, m] for l in range(0, 3))
              for i in range(0, 3) for m in range(0, 12)), name='Transport units')

# 6. Equal wear and tear
m.addConstrs((s[i, 0, m] / w[i, 0] == s[i, 1, m] / w[i, 1] 
              for i in range(0, 3) for m in range(0, 12)), name = '1')

m.addConstrs((s[i, 1, m] / w[i, 1] == s[i, 2, m] / w[i, 2] 
              for i in range(0, 3) for m in range(0, 12)), name = '2')

m.addConstrs((s[i, 0, m] / w[i, 0] == s[i, 2, m] / w[i, 2] 
              for i in range(0, 3) for m in range(0, 12)), name = '3')

# m.addConstrs((s[0, l, m] / w[0, l] == s[1, l, m] / w[1, l] for l in range(0, 3) for m in range(0, 12)), name = 'Each factory all same 1')
# m.addConstrs((s[1, l, m] / w[1, l] == s[2, l, m] / w[2, l] for l in range(0, 3) for m in range(0, 12)), name = 'Each factory all same 2')

# Run optimization
m.optimize()


obj = m.getObjective()
model2_price = obj.getValue()
model2_price



# Output
factory = []
production_line = []
month = []
quantity = []
for facility in s.keys():
    factory.append(Factory[facility[0]])
    production_line.append(Line[facility[1]])
    month.append(Month[facility[2]])
    quantity.append(s[facility].x)

model2 = pd.DataFrame({'Factory':factory, 'ProductionLine':production_line,'Month':month, 'Quantity':quantity})
model2.to_excel('Model2_dif_quantity.xlsx', index = 0)



# # Model 3



m = gp.Model("model 3")

# Transportation cost
# shipping_cost
ship = pd.read_excel('Project Data.xlsx', 'shipping_cost')
ship = ship.iloc[:, 1:]
c = {(facility, customer): ship.iloc[facility, customer]
            for facility in range(0, 3)
            for customer in range(0, 8)}
# Production rate of production line l at factory i (Units per hour) 
w = {(facility, line): 6
            for facility in range(0, 3)
            for line in range(0, 3)}
# Production capacity : M

M = {(facility, line): 6 * 160
            for facility in range(0, 3)
            for line in range(0, 3)}
# energy cost
energy = pd.read_excel('Project Data.xlsx', 'energy_cost_new')
energy = energy.iloc[:, 1:]
e = {(facility, line): energy.iloc[0,line]
            for facility in range(0, 3)
            for line in range(0, 3)}
# demand
demand = pd.read_excel('Project Data.xlsx', 'demand')
demand = demand.iloc[:, 1:]
d = {(month, customer): demand.iloc[month, customer]
            for month in range(0, 12)
            for customer in range(0, 8)}


# ### Decision variable

# the number of units transported from factory i to store j in month m
u = m.addVars(list(product(range(0, 3), range(0, 8), range(0, 12))),lb = 0 ,vtype = GRB.INTEGER, name = 'units transported')

# the number of units produced at factory i from production line l in month m
s = m.addVars(list(product(range(0, 3), range(0, 3), range(0, 12))),lb = 0,vtype = GRB.INTEGER, name = 'units produced')

# Inventory of store j in month m
ii = m.addVars(list(product(range(0, 8), range(0, 12))),lb = 0,vtype = GRB.INTEGER, name = 'inventory')


# ### Objective function

m.setObjective((3 * gp.quicksum(1.2 * c[i,j] * u[i, j, m] + e[i, l] * s[i, l, m] for l in range(0, 3) 
                             for m in range(0, 12) for j in range(0, 8) for i in range(0, 3)) 
               ), GRB.MINIMIZE)


# ### Constractions

# 1. Demand Fulfilment on Transportation
m.addConstrs((gp.quicksum(u[i, j, m] for i in range(0, 3)) >= d[m, j] 
              for j in range(0, 8) for m in range(0, 12)), name='Demand')

# 2. Production Capacity
m.addConstrs((s[i, l, m] <= M[i, l] 
              for l in range(0, 3) for i in range(0, 3) for m in range(0, 12)), name='Capacity')

# 3. Demand Fulfilment on Factory
m.addConstrs((gp.quicksum(s[i, l, m] for i in range(0, 3) for l in range(0, 3)) >= (gp.quicksum(d[m, j] 
            for j in range(0, 8))) for m in range(0, 12)), name='Demand')

# 4. Transportation units can’t exceed production units
m.addConstrs((gp.quicksum(u[i, j, m] for j in range(0, 8)) <= gp.quicksum(s[i, l, m] for l in range(0, 3))
              for i in range(0, 3) for m in range(0, 12)), name='Transport units')

# 6. Equal wear and tear
m.addConstrs((s[i, 0, m] / w[i, 0] == s[i, 1, m] / w[i, 1] 
              for i in range(0, 3) for m in range(0, 12)), name = '1')

m.addConstrs((s[i, 1, m] / w[i, 1] == s[i, 2, m] / w[i, 2] 
              for i in range(0, 3) for m in range(0, 12)), name = '2')

m.addConstrs((s[i, 0, m] / w[i, 0] == s[i, 2, m] / w[i, 2] 
              for i in range(0, 3) for m in range(0, 12)), name = '3')

# m.addConstrs((s[0, l, m] / w[0, l] == s[1, l, m] / w[1, l] for l in range(0, 3) for m in range(0, 12)), name = 'Each factory all same 1')
# m.addConstrs((s[1, l, m] / w[1, l] == s[2, l, m] / w[2, l] for l in range(0, 3) for m in range(0, 12)), name = 'Each factory all same 2')

# Run optimization
m.optimize()


obj = m.getObjective()
model3_price = obj.getValue()

# Output
factory = []
production_line = []
month = []
quantity = []
for facility in s.keys():
    factory.append(Factory[facility[0]])
    production_line.append(Line[facility[1]])
    month.append(Month[facility[2]])
    quantity.append(s[facility].x)

model3 = pd.DataFrame({'Factory':factory, 'ProductionLine':production_line,'Month':month, 'Quantity':quantity})
model3.to_excel('Model3_quantity.xlsx', index = 0)


# Reserve price

3 * model1_price - model3_price


# ## Q1

print(model1_price - model2_price)


# ### Q2

print(3 * model1_price - model3_price)

