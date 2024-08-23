from gurobipy import *
import openpyxl as oxl
from openpyxl.styles import Alignment
import numpy as np
import pandas as pd

model = Model('optimizer')

def solve(I_total,              # set of all customers and Medical center
          I0,                   # set of all patients pick up
          I_0,                  # set of all patients drop off
          I1,                   # set of all clients with staff 1 requirement drop off node
          I_1,                  # set of all clients with staff 1 requirement pick up node
          I2,                   # set of all clients with staff 2 requirement
          I_2,                  # set of all clients with staff 2 requirement pick up node
          I3,                   # set of all clients with staff 3 requirement
          I_3,                  # set of all clients with staff 3 requirement pick up node
          tt,                   # set of travel times (call the travel time from i to j by punching in "tt[j].get(i)"
          EST_patient, 
          LST_patient, 
          STD_patient, 
          EST_client, 
          LST_client, 
          STD_client,
          S,                    # set of all staff
          S1,                   # set of all staff level 1
          S2,                   # set of all staff level 2
          S3,                   # set of all staff level 3
          vehicle_capacity,     # vehicle_capacity
          number_of_vehicles,   # how many vehicles are available
          allowed_routes,       # how many different routes are allowed
          allow_wait_staff,     # is staff allowed to wait at clients
          allow_delay_clients,  # is delay at clients allowed
          allow_wait_patients,   # are patients allowed to wait
          allow_overtime,       # is staff allowed to work overtime
          max_wait_staff,       # how long is staff allowed to wait at most
          max_wait_patiens,     # how long are patiens allowed to wait at most
          level_1,              # shift cost for staff level 1
          level_2,              # shift cost for staff level 2
          level_3,              # shift cost for staff level 3
          over_time,            # additional wage surplus for overtime
          driver,               # shift cost for driver
          car_fixed,            # fixed cost for a car being used
          fuel_cost,            # fuel cost per minute
          max_driver,           # maximum time a driver is allowed to drive
          ):
    
    # node_pairs_1 = []
    # node_pairs_2 = []
    # node_pairs_3 = []
    # for i in range(len(I1)):
    #     node_pairs_1 += [(I1[i], I_1[i])]
    # for i in range(len(I2)):
    #     node_pairs_2 += [(I2[i], I_2[i])]
    # for i in range(len(I3)):
    #     node_pairs_3 += [(I3[i], I_3[i])]

    # all_node_pairs = node_pairs_1 + node_pairs_2 + node_pairs_3 + [('MC', 'MCd')]

    node_staff_1 = [] # combinations of i and s with qualification 1 list of Tuples with i as the node and s as the allowed staff
    node_staff_2 = [] # combinations of i and s with qualification 2
    node_staff_3 = [] # combinations of i and s with qualification 3
    
    for i in (I1 + ['MCd']):
        for s in S1:
            node_staff_1 += [(i,s)]
    for i in (I2 + ['MCd']):
        for s in S2:
            node_staff_2 += [(i,s)]
    for i in (I3 + ['MCd']):
        for s in S3:
            node_staff_3 += [(i,s)]
    
    node_staff_tuple = node_staff_1 + node_staff_2 + node_staff_3

    I_dict = {}
    I0_dict = {}
    for i in range(len(I0)):
        I0_dict[I0[i]] = I_0[i]
        I_dict[I0[i]] = I_0[i]
    
    I1_dict = {}
    for i in range(len(I1)):
        I1_dict[I1[i]] = I_1[i]
        I_dict[I1[i]] = I_1[i]
    
    I2_dict = {}
    for i in range(len(I2)):
        I2_dict[I2[i]] = I_2[i]
        I_dict[I2[i]] = I_2[i]
    
    I3_dict = {}
    for i in range(len(I3)):
        I3_dict[I3[i]] = I_3[i]
        I_dict[I3[i]] = I_3[i]
    
    I_dict['MCd'] = 'MC'
   

    M = 1000

    """Variable Declaration"""

    X = {}
    t = {}
    W = {}
    D = {}
    P = {}
    Dt = {}
    Pt = {}
    p = {}
    d = {}
    pt = {}
    dt = {}
    y = {}
    o = {}
    op = {}
    wait_at_client_before = {}
    wait_at_client_after = {}
    delay_at_client = {}
    wait_at_MC_before = {}
    wait_at_MC_after = {}
    use_car = {}
    route_car_match = {}
    predecessor = {}
    start_car_travel = {}
    end_car_travel = {}
    car_travel = {}
    use_staff = {}
    overtime = {}
    
    for i in I_total:
        for r in range(allowed_routes):
            for j in I_total:
                if i != j:
                    # does route r visit node j after node i (=1) or not (=0):
                    X[i,j,r] = model.addVar(name=f'X_{i}_{j}_{r}', vtype='b') 
                if (i != j) & (i != 'MCd'):
                    for s in S:    
                        # is staff s on route r while going from i to j (=1) or not (=0):
                        o[i,j,s,r] = model.addVar(name=f'o_{i}_{j}_{s}_{r}', vtype='b')
                    for l in I_0:
                        # is patient l riding on route r from node i to node j (=1) or not (=0):
                        op[i,j,l,r] = model.addVar(name=f'op_{i}_{j}_{l}_{r}', vtype='b')  
            # at which time in the route is the route r arriving at node i:
            t[i,r] = model.addVar(name=f't_{i}_{r}', vtype='c', lb=0)
    
    for (i,s) in node_staff_tuple:
        # is staff s performing task at client i of appropriate level (=1) or not (=0):
        W[i,s] = model.addVar(name=f'W_{i}_{s}', vtype='b') 
        # at what time is staff s dropped at node i:
        Dt[i,s] = model.addVar(name=f'Dt_{i}_{s}', vtype='c', lb=0) 
        # at what time is staff s picked up from node ip:
        Pt[I_dict.get(i),s] = model.addVar(name=f'Pt_{I_dict.get(i)}_{s}', vtype='c', lb=0) 
        # how long has staff s to wait at client i before beginning service:
        wait_at_client_before[i,s] = model.addVar(name=f'wait_at_client_before_{s}_{i}', vtype='c', lb=0)
        wait_at_client_after[i,s] = model.addVar(name=f'wait_at_client_after_{s}_{i}', vtype='c', lb=0)
        # how long is the delay of staff s at client i for service start:
        delay_at_client[i,s] = model.addVar(name=f'service_delay_{i}_{s}', vtype='c', lb=0) 
        for r in range(allowed_routes): 
            # does route r drop staff s off at node i (=1) or not (=0):
            D[i,s,r] = model.addVar(name=f'D_{i}_{s}_{r}', vtype='b')
            # does route r pick staff s up from node ip (=1) or not (=0):
            P[I_dict.get(i),s,r] = model.addVar(name=f'P_{I_dict.get(i)}_{s}_{r}', vtype='b') 

    for r in range(allowed_routes):
        for r2 in range(allowed_routes):
            if r != r2:
                # does route r start before route r2 (=1) or not (=0):
                predecessor[r,r2] = model.addVar(name=f'predecessor_{r}_{r2}', vtype='b')             
    
    for i in I0:
        # at what time is patient i dropped off at MC:
        dt[i] = model.addVar(name=f'dt_{i}', vtype='c', lb=0) 
        # at what time is i picked up from MCd:
        pt[I_dict.get(i)] = model.addVar(name=f'pt_{I_dict.get(i)}', vtype='c', lb=0) 
        for r in range(allowed_routes): 
            # is route r dropping off patient i at their drop off node ip (=1) or not (=0):
            d[I_dict.get(i),r] = model.addVar(name=f'd_{I_dict.get(i)}_{r}', vtype='b') 
            # is route r picking up patient i at their pickup node i (=1) or not (=0):
            p[i,r] = model.addVar(name=f'p_{i}_{r}', vtype='b') 
        # how long does patient l have to wait at MC:
        wait_at_MC_before[i] = model.addVar(name=f'wait_at_MC_before_{i}', vtype='c', lb=0)
        wait_at_MC_after[i] = model.addVar(name=f'wait_at_MC_after_{i}', vtype='c', lb=0)
    
    for i in (I1 + I2 + I3):
            # how long is the delay of staff s at client i for service start:
            delay_at_client[i] = model.addVar(name=f'service_delay_{i}', vtype='c', lb=0) 

    for c in range(number_of_vehicles):
        # is car c used (=1) or not (=0):
        use_car[c] = model.addVar(name=f'use_car_{c}', vtype='b') 
        # at what time does car c start its travel for the day:
        start_car_travel[c] = model.addVar(name=f'start_car_travel_{c}', vtype='c', lb=0)
        # at what time does car c end its travel for the day: 
        end_car_travel[c] = model.addVar(name=f'end_car_travel_{c}', vtype='c', lb=0)
        # what is the operating time of car c for the day:
        car_travel[c] = model.addVar(name=f'car_travel_{c}', vtype='c', lb=0)
        for r in range(allowed_routes):
            # does car c drive route r (=1) or not (=0):
            route_car_match[c,r] = model.addVar(name=f'route_{r}_done_by_car_{c}', vtype='b') 

    # is staff s scheduled or not:
    for s in S:
        use_staff[s] = model.addVar(name=f'use_staff_{s}')
        overtime[s] = model.addVar(name=f'overtime_{s}', vtype='c', lb=0, ub=120)