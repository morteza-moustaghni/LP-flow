import xlrd, xlwt
from gurobipy import *

#####################################################	PART 1	###########################################################

book = xlrd.open_workbook('input2.xlsx')
wb_terminaler = book.sheet_by_name('Terminaler')
wb_destruktion = book.sheet_by_name('Destruktionspl')
wb_prod_kapacitet = book.sheet_by_name('Produktionskapacitet')
wb_fabriker = book.sheet_by_name('Fabriker')
wb_produkter = book.sheet_by_name('Produkter')
wb_dist_dest = book.sheet_by_name('Distans2')
wb_distribution = book.sheet_by_name('Distributionskostnad')
wb_behov = book.sheet_by_name('Behov')
wb_regioner = book.sheet_by_name('Regioner')
wb_fabrik_dist = book.sheet_by_name('Distans1')

p_rodukter = [] # p
f_abriker = [] # i
d_istributionsterminaler = [] # j
r_egioner = [] # k
d_estruktionsterminaler = [] # l

behov_per_region = {}												# Demand per region
kapacitet_fabrik_produkt = {}										# Capacity for each factory of each product
kostnad_fabrik_dist = {}											# Cost of transportation between factories and distribution terminals
kostnad_dist_region = {}											# Cost of transportation between distribution terminals and customers
kostnad_dist_dest = {}												# Cost of transportation between distribution terminals to destruction terminals

for i in range(1, wb_produkter.nrows): 								# Products
	p_rodukter.append(wb_produkter.cell_value(i, 0).strip())
for i in range(1, wb_fabriker.nrows): 								# Factories
	f_abriker.append(wb_fabriker.cell_value(i, 0).strip())
for i in range(1, wb_terminaler.nrows): 							# Distribution Terminals
	d_istributionsterminaler.append(wb_terminaler.cell_value(i, 0).strip())
for i in range(1, wb_regioner.nrows):								# Regions
	r_egioner.append(wb_regioner.cell_value(i, 0).strip())
for i in range(1, wb_destruktion.nrows):							# Destruction Terminals
	d_estruktionsterminaler.append(wb_destruktion.cell_value(i, 0).strip())

i = 1
while i < wb_behov.nrows: # DEMAND OF EACH PRODUCT AT EACH REGION
	key = (wb_behov.cell_value(i, 0), wb_behov.cell_value(i, 1).strip())
	value = wb_behov.cell_value(i, 2)
	behov_per_region[(key)] = value
	i += 1

i = 1
while i < wb_prod_kapacitet.nrows: # CAPACITY OF EACH FACTORY OF EACH PRODUCT
	key = (wb_prod_kapacitet.cell_value(i, 0), wb_prod_kapacitet.cell_value(i, 1).strip())
	value = wb_prod_kapacitet.cell_value(i, 2)
	kapacitet_fabrik_produkt[(key)] = value
	i += 1

i = 1
while i < wb_fabrik_dist.nrows: # COST OF TRANSPORTATION, FACTORY TO DISTRIBUTION TERMINAL
	key = (wb_fabrik_dist.cell_value(i, 0), wb_fabrik_dist.cell_value(i, 1))
	value = wb_fabrik_dist.cell_value(i, 2) * wb_fabrik_dist.cell_value(i, 3)
	kostnad_fabrik_dist[(key)] = value
	i += 1

i = 1
while i < wb_dist_dest.nrows: # COST OF TRANSPORTATION, DISTRIBUTION TERMINAL TO DESTRUCTION TERMINAL
	key = (wb_dist_dest.cell_value(i, 0), wb_dist_dest.cell_value(i, 1))
	value = wb_dist_dest.cell_value(i, 2) * wb_dist_dest.cell_value(i, 3)
	kostnad_dist_dest[(key)] = value
	i += 1

i = 1
while i < wb_distribution.nrows: # COST OF TRANSPORTATION, DISTRIBUTION TERMINAL TO CUSTOMER
	key = (wb_distribution.cell_value(i, 0), wb_distribution.cell_value(i, 1))
	value = wb_distribution.cell_value(i, 2)
	kostnad_dist_region[(key)] = value
	i += 1

#######################################################################	PART 2	##############################################################
# Gurobi optimization

gurobimodel = Model()

x_ijd = {}
x_jid = {}
x_jkd = {} #Triggar y_jk
x_jld = {}

y_jk = {}
f_j = {}
a_l = {}

# Making new gurobi variables for flow to and from distribution terminals from and to factories of each product
for i in f_abriker:
	for j in d_istributionsterminaler:
		for d in p_rodukter:
			x_ijd[(i, j, d)] = gurobimodel.addVar(lb=0, ub=GRB.INFINITY, vtype=GRB.CONTINUOUS)
			x_jid[(j, i, d)] = gurobimodel.addVar(lb=0, ub=GRB.INFINITY, vtype=GRB.CONTINUOUS)

for j in d_istributionsterminaler:
	for k in r_egioner:
		key = (j, k)
		y_jk[(key)] = gurobimodel.addVar(lb=0, ub=1, vtype=GRB.BINARY)
		for d in p_rodukter:
			x_jkd[(j, k, d)] = gurobimodel.addVar(lb=0, ub=GRB.INFINITY, vtype=GRB.CONTINUOUS)

for j in d_istributionsterminaler:
	for l in d_estruktionsterminaler:
		for d in p_rodukter:
			key = (j, l, d)
			x_jld[(key)] = gurobimodel.addVar(lb=0, ub=GRB.INFINITY, vtype=GRB.CONTINUOUS)

for j in d_istributionsterminaler:
	key = j
	value = gurobimodel.addVar(lb=0, ub=1, vtype=GRB.BINARY)
	f_j[(key)] = value

for l in d_estruktionsterminaler:
	key = l
	value = gurobimodel.addVar(lb=0, ub=1, vtype=GRB.BINARY)
	a_l[(key)] = value

# Constraints
# 7
for i in f_abriker:
	for d in p_rodukter:
		gurobimodel.addConstr(quicksum(x_ijd[i, j, d] for j in d_istributionsterminaler) <= kapacitet_fabrik_produkt[i, d])

# HERE WE MAKE SURE OLD PRODUCTS ARE ONLY SENT TO FACTORIES THAT CAN REPRODUCE THEM. ASSUMING THIS, WE CAN EXPECT THE PRODUCTION COST OF THOSE TO BE HALVED.
for i in f_abriker:
	for d in p_rodukter:
		gurobimodel.addConstr(quicksum(x_jid[j, i, d] for j in d_istributionsterminaler) <= kapacitet_fabrik_produkt[i, d])

# 8
for j in d_istributionsterminaler:
	for d in p_rodukter:
		gurobimodel.addConstr(0.6*quicksum(x_ijd[i,j,d] for i in f_abriker) - quicksum(x_jid[j,i,d] for i in f_abriker ) == 0)		

# 9
for j in d_istributionsterminaler:
	for d in p_rodukter:
		gurobimodel.addConstr(0.4*quicksum(x_ijd[i,j,d] for i in f_abriker) - quicksum(x_jld[j,l,d] for l in d_estruktionsterminaler) == 0)

# SLBD
M = 1000000000

# 10
for j in d_istributionsterminaler:
	for k in r_egioner:
		gurobimodel.addConstr(quicksum(x_jkd[(j, k, d)] for d in p_rodukter) <= y_jk[(j, k)] * M)

# 11
for j in d_istributionsterminaler:
	gurobimodel.addConstr(quicksum(x_ijd[(i, j, d)] for i in f_abriker for d in p_rodukter) <= f_j[(j)] * M)

# 12
for j in d_istributionsterminaler:
	gurobimodel.addConstr(quicksum(x_jid[(j, i, d)] for i in f_abriker for d in p_rodukter) <= f_j[(j)] * M)

# 13
for j in d_istributionsterminaler:
	gurobimodel.addConstr(quicksum(x_jld[(j, l, d)] for l in d_estruktionsterminaler for d in p_rodukter) <= f_j[(j)] * M)

# 14
for j in d_istributionsterminaler:
	for l in d_estruktionsterminaler:
		gurobimodel.addConstr(quicksum(x_jld[(j, l, d)] for d in p_rodukter) <= a_l[(l)] * M)

# 15
for k in r_egioner:
	gurobimodel.addConstr(quicksum(y_jk[(j, k)] for j in d_istributionsterminaler) == 1)

# 16
for j in d_istributionsterminaler:
	for d in p_rodukter:
		gurobimodel.addConstr(quicksum(x_ijd[i,j,d] for i in f_abriker) - quicksum(x_jkd[j, k, d] for k in r_egioner) == 0)

# 17
for k in r_egioner:
	for d in p_rodukter:
		gurobimodel.addConstr(quicksum(x_jkd[(j, k, d)] for j in d_istributionsterminaler) - behov_per_region[(k, d)] == 0)

# MALFUNKTION:
f_to_di = LinExpr(quicksum(kostnad_fabrik_dist[i, j] * x_ijd[(i, j, d)] for i in f_abriker for j in d_istributionsterminaler for d in p_rodukter))

di_to_re = LinExpr(quicksum(kostnad_dist_region[j, k] * y_jk[(j, k)] for j in d_istributionsterminaler for k in r_egioner))

di_to_de = LinExpr(quicksum(kostnad_dist_dest[j, l] * x_jld[(j, l, d)] for j in d_istributionsterminaler for l in d_estruktionsterminaler for d in p_rodukter))

di_to_f = LinExpr(quicksum(kostnad_fabrik_dist[i, j] * x_jid[(j, i, d)] for i in f_abriker for j in d_istributionsterminaler for d in p_rodukter))

f_cost_di = LinExpr(quicksum(1000000 * f_j[(j)] for j in d_istributionsterminaler))

f_cost_de = LinExpr(quicksum(50000 * a_l[(l)] for l in d_estruktionsterminaler))

p_roduction_cost = LinExpr(quicksum(x_ijd[(i, j, d)] for i in f_abriker for j in d_istributionsterminaler for d in p_rodukter))

prod_cost_old = LinExpr(-0.5*quicksum(x_jid[(j, i, d)] for j in d_istributionsterminaler for i in f_abriker for d in p_rodukter))

gurobimodel.setObjective(f_to_di + di_to_re + di_to_de + di_to_f + f_cost_di + f_cost_de + p_roduction_cost + prod_cost_old, GRB.MINIMIZE)
gurobimodel.optimize()

############################################################# Part 3 #######################################################
# Validation

validering = xlwt.Workbook()
con_j_k = validering.add_sheet('con_j_k')
flo_j_k_d = validering.add_sheet('flo_j_k_d')
flo_i_j_d = validering.add_sheet('flo_i_j_d')
flo_j_l_d = validering.add_sheet('flo_j_l_d')
flo_j_i_d = validering.add_sheet('flo_j_i_d')
num_xg = validering.add_sheet('num_xg')
num_xn = validering.add_sheet('num_xn')

a = 0
for j in d_istributionsterminaler:
	for k in r_egioner:
		key = (j,k)
		if y_jk[j, k].x == 1:
			con_j_k.write(a, 0, j)
			con_j_k.write(a, 1, k)
			a += 1

a = 0
for j in d_istributionsterminaler:
	for k in r_egioner:
		for d in p_rodukter:
			key = (j,k,d)
			if x_jkd[(key)].x > 0:
				flo_j_k_d.write(a, 0, j)
				flo_j_k_d.write(a, 1, k)				
				flo_j_k_d.write(a, 2, d)
				flo_j_k_d.write(a, 3, x_jkd[(key)].x)
				a += 1 

a = 0
for i in f_abriker:
	for j in d_istributionsterminaler:
		for d in p_rodukter:
			key = (i,j,d)
			if x_ijd[(key)].x > 0:
				flo_i_j_d.write(a, 0, i)
				flo_i_j_d.write(a, 1, j)				
				flo_i_j_d.write(a, 2, d)
				flo_i_j_d.write(a, 3, x_ijd[(key)].x)
				a += 1

a = 0
for j in d_istributionsterminaler:
	for l in d_estruktionsterminaler:
		for d in p_rodukter:
			key = (j,l,d)
			if x_jld[(key)].x > 0:
				flo_j_l_d.write(a, 0, j)
				flo_j_l_d.write(a, 1, l)				
				flo_j_l_d.write(a, 2, d)
				flo_j_l_d.write(a, 3, x_jld[(key)].x)
				a += 1

a = 0
for j in d_istributionsterminaler:
	for i in f_abriker:
		for d in p_rodukter:
			key = (j,i,d)
			if x_jid[(key)].x > 0:
				flo_j_i_d.write(a, 0, j)
				flo_j_i_d.write(a, 1, i)				
				flo_j_i_d.write(a, 2, d)
				flo_j_i_d.write(a, 3, x_jid[(key)].x)
				a += 1

final = validering.add_sheet('Final')
final.write(0, 0, gurobimodel.getAttr(GRB.Attr.ObjVal))
validering.save('validering.xls')