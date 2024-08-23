"""Unified Modeling Approach"""

from data import get_data
from optimizer import solve
from plot_graph import plotter
from plot_solverlog import log_plotter
from gurobipy import *
import glob
import os

"""Delete all Output Files"""

files = glob.glob(os.path.join('outputs', "*"))

# Iterate over the list and delete each file
for file in files:
    try:
        os.remove(file)
        print(f"Deleted file: {file}")
    except Exception as e:
        print(f"Error deleting file {file}: {e}")


"""Solver Settings"""
runtime_limit = 172800      # in seconds
MIPfocus = 1                # 1 = find feasible solutions; 2 = proof optimality; 3 = improve obj bound
heuristics = 0.5            # between 0 and 1 fraction of time that is used for heuristics


"""Parameter Setting Section"""

#data set size:
data_size = 10

# available nurses (staff):
staff1 = 1
staff2 = 1
staff3 = 1

# number of vehicles and vehicle capacity:
vehicle_capacity = 4
number_of_vehicles = 1

# set of Booleans that are used to turn on and off specific constraints:
allow_wait_staff = False
allow_delay_clients = False
allow_wait_patients = False
allow_overtime = False
max_wait_staff = 60
max_wait_patients = 60

# staff cost per qualification, overtime wage, driver cost, driver overtime wage, fixed cost car, fuel cost car:
level_1 = 200   # 25€/h
level_2 = 240   # 30€/h
level_3 = 280   # 35€/h
over_time = 20  # 20€/h overtime surplus 
driver = 160    # 20€/h
car_fixed = 50  # daily cost of car
fuel_cost = 0
# fuel_cost = 25/60 * 8/100 * 1.8 # 25 km/h average speed --> 25/60 km/min, 8 l / 100 km fuel economy in city traffic, 1.8 € per liter average price
driver_time_limit = 480 # driver shift limit in minutes



S1 = []
S2 = []
S3 = []
for s in range(staff1):
    S1 += [f's{s+1}']
for s in range(staff2):
    S2 += [f's{staff1 + s+1}']
for s in range(staff3):
    S3 += [f's{staff1 + staff2 + s+1}']
S = S1 + S2 + S3
allowed_routes = 10 * number_of_vehicles 

output_file_path = "outputs/a_input_parameters.txt"
with open(output_file_path, "w") as output_file:
    output_file.write(f"Data size: {data_size} \n \n")
    output_file.write(f"Costs: \n Staff level 1: {level_1} \n Staff level 2: {level_2} \n Staff level 3: {level_3} \n Overtime: {over_time} \n Car fixed: {car_fixed} \n Fuel cost: {fuel_cost} \n \n")
    output_file.write(f"Solution Space Booleans: \n Allow wait staff: {allow_wait_staff} \n Allow wait patients: {allow_wait_patients} \n Allow delay clients: {allow_delay_clients} \n Allow overtime: {allow_overtime} \n \n")
    output_file.write(f"Other parameters: \n Staff 1: {staff1} \n Staff 2: {staff2} \n Staff 3: {staff3} \n Vehicle capacity: {vehicle_capacity} \n Number of vehicles: {number_of_vehicles} \n Number of allowed routes: {allowed_routes} \n \n")
    output_file.write(f"Solver Settings: \n Runtime limit: {runtime_limit} \n MIPfocus: {MIPfocus} \n Fraction for heuristics: {heuristics}")


# import data:
I_total, I0, I_0, I1, I_1, I2, I_2, I3, I_3, tt, EST_patient, LST_patient, STD_patient, EST_client, LST_client, STD_client= get_data(f'data/data_{data_size}.xlsx')

# solve model:
try:
    print(f'Data Set Size: {data_size}')
    model_status, arcs, node_info = solve(I_total, I0, I_0, I1, I_1, I2, I_2, I3, I_3, tt, EST_patient, LST_patient, STD_patient, EST_client, LST_client, STD_client, S, S1, S2, S3, 
                                        vehicle_capacity, 
                                        number_of_vehicles, 
                                        allowed_routes,
                                        allow_wait_staff,
                                        allow_delay_clients,
                                        allow_wait_patients,
                                        allow_overtime,
                                        max_wait_staff,
                                        max_wait_patients,
                                        level_1, level_2, level_3, over_time, driver, car_fixed, fuel_cost,
                                        driver_time_limit,
                                        runtime_limit,
                                        MIPfocus,
                                        heuristics
                                        )
    # if optimal, plot the resulting graph:
    if model_status == GRB.OPTIMAL:
        plotter(arcs, node_info, I_total, allowed_routes)
except TypeError as e:
    print("Caught TypeError in solve function:", e)

# if model_status == GRB.TIME_LIMIT or model_status == GRB.OPTIMAL:
#     log_plotter(log_callback)