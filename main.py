from data import get_data
from routes import solve
from plot_graph import plotter
from plot_solverlog import log_plotter
from gurobipy import *

# set basic boundaries of the system, such as vehicle capacities and numer of available vehicles:
vehicle_capacity = 4
number_of_vehicles = 2
# here seems to be a lot of computational complexity introduced --> I think this causes symmetry because the routes are not sorted by time
# incresing from 2 vehicles with 4 total routes to 2 vehicles with 8 total routes increases solver time by factor of about 40 
allowed_routes = 2 * number_of_vehicles 
# set of Booleans that are used to turn on and off specific constraints:
allow_wait_staff = False
allow_delay_clients = False
allow_wait_patients = False
max_wait_staff = 60
max_wait_patients = 60

# available nurses (staff):
S1 = ['s1',]
S2 = ['s2',]
S3 = ['s3',]
# S1 = ['s1','s4',]
# S2 = ['s2','s5',]
# S3 = ['s3','s6',]
S = S1 + S2 + S3

# import data:
I_total, I0, I_0, I1, I_1, I2, I_2, I3, I_3, tt, EST_patient, LST_patient, STD_patient, EST_client, LST_client, STD_client= get_data('data/data_10.xlsx')

# solve model:
try:
    model_sol, arcs, node_info = solve(I_total, I0, I_0, I1, I_1, I2, I_2, I3, I_3, tt, EST_patient, LST_patient, STD_patient, EST_client, LST_client, STD_client, S, S1, S2, S3, 
                                        vehicle_capacity, 
                                        number_of_vehicles, 
                                        allowed_routes,
                                        allow_wait_staff,
                                        allow_delay_clients,
                                        allow_wait_patients,
                                        max_wait_staff,
                                        max_wait_patients,
                                        )
    # if optimal, plot the resulting graph:
    if model_sol > 0:
        plotter(arcs, node_info, I_total, allowed_routes)
except TypeError as e:
    print("Caught TypeError in solve function:", e)

# if model_status == GRB.TIME_LIMIT or model_status == GRB.OPTIMAL:
#     log_plotter(log_callback)