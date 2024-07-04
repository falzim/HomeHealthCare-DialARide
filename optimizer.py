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
          fuel_cost             # fuel cost per minute
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
        for j in I_total:
            if i != j:
                for r in range(allowed_routes):
                    # does route r visit node j after node i (=1) or not (=0):
                    X[i,j,r] = model.addVar(name=f'X_{i}_{j}_{r}', vtype='b') 

    for i in I_total: 
        for r in range(allowed_routes):
            # at which time in the route is the route r arriving at node i:
            t[i,r] = model.addVar(name=f't_{i}_{r}', vtype='c', lb=0) 

    # t = model.addVars(I_total, range(allowed_routes), name='t', vtype='b')
    
    for (i,s) in node_staff_tuple:
        # is staff s performing task at client i of appropriate level (=1) or not (=0):
        W[i,s] = model.addVar(name=f'W_{i}_{s}', vtype='b') 
        # at what time is staff s dropped at node i:
        Dt[i,s] = model.addVar(name=f'Dt_{i}_{s}', vtype='c', lb=0) 
        # at what time is staff s picked up from node ip:
        Pt[I_dict.get(i),s] = model.addVar(name=f'Pt_{I_dict.get(i)}_{s}', vtype='c', lb=0) 
        # if allow_wait_staff:
        #     # how long has staff s to wait at client i before beginning service:
        #     wait_at_client_before[i,s] = model.addVar(name=f'wait_at_client_before_{s}_{i}', vtype='c', lb=0)
        #     wait_at_client_after[i,s] = model.addVar(name=f'wait_at_client_after_{s}_{i}', vtype='c', lb=0)
        # # how long is the delay of staff s at client i for service start:
        # delay_at_client[i,s] = model.addVar(name=f'service_delay_{i}_{s}', vtype='c', lb=0) 
        for r in range(allowed_routes): 
            # does route r drop staff s off at node i (=1) or not (=0):
            D[i,s,r] = model.addVar(name=f'D_{i}_{s}_{r}', vtype='b')
            # does route r pick staff s up from node ip (=1) or not (=0):
            P[I_dict.get(i),s,r] = model.addVar(name=f'P_{I_dict.get(i)}_{s}_{r}', vtype='b') 

    for r in range(allowed_routes):
        for i in I_total:
            # is route r going to node i (=1) or not (=0):
            y[i,r] = model.addVar(name=f'y_{i}_{r}', vtype='b') 
        for r2 in range(allowed_routes):
            if r != r2:
                # does route r start before route r2 (=1) or not (=0):
                predecessor[r,r2] = model.addVar(name=f'predecessor_{r}_{r2}', vtype='b')             
    
    for s in S:
        for r in range(allowed_routes):
            for i in I_total:
                for j in I_total:
                    if (i != j) & (i != 'MCd'):
                        # is staff s on route r while going from i to j (=1) or not (=0):
                        o[i,j,s,r] = model.addVar(name=f'o_{i}_{j}_{s}_{r}', vtype='b') 
    
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
    
    
    for l in I0:
        # if allow_wait_patients:
        #     # how long does patient l have to wait at MC:
        #     wait_at_MC_before[l] = model.addVar(name=f'wait_at_MC_before_{l}', vtype='c', lb=0)
        #     wait_at_MC_after[l] = model.addVar(name=f'wait_at_MC_after_{l}', vtype='c', lb=0)
        for i in I_total:
            for j in I_total:
                if (i != j):
                    for r in range(allowed_routes):
                        # is patient l riding on route r from node i to node j (=1) or not (=0):
                        op[i,j,l,r] = model.addVar(name=f'op_{i}_{j}_{l}_{r}', vtype='b') 

    # if allow_delay_clients:
    #     for i in (I1 + I2 + I3):
    #         # how long is the delay of staff s at client i for service start:
    #         delay_at_client[i] = model.addVar(name=f'service_delay_{i}', vtype='c', lb=0) 

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

    if allow_overtime:
        for s in S: 
            overtime[s] = model.addVar(name=f'overtime_{s}', vtype='c', lb=0, ub=120)
    
    
    """---------------------------------------------------------------------------------------------------------------------"""
    """---------------------------------------------------------------------------------------------------------------------"""
    """---------------------------------------------------------------------------------------------------------------------"""
    """Setting Constraints"""
    

    """Flow Conservation, Making sure every node is visited and SEC"""
    # for r in range(allowed_routes):
    #     z['MC', r] = 0 # set MC as tour starting point

    # for r in range(allowed_routes):
    #     for i in I_total:
    #         for j in I_total:
    #             if i != j:
    #                 # subtour elimination constraint MTZ:
    #                 # model.addConstr(z[i,r] - z[j,r] + len(I_total) * X[i,j,r] <= len(I_total) - 1)
    #                 model.addConstr(z[i,r] + 1 <= z[j,r] + (1 - X[i,j,r]) * len(I_total))
    #                 ## is this actually necessary? Time windows could be fulfilling the same 
    
    for r in range(allowed_routes):
        # no incoming arcs to the depot start node
        model.addConstr(quicksum(X[j,'MC',r] for j in I_total if j != 'MC') == 0)
        # no outgoing arcs from depot end node
        model.addConstr(quicksum(X['MCd',j,r] for j in I_total if j != 'MCd') == 0)
        # exactly one outgoing arc from depot start node
        model.addConstr(quicksum(X['MC',j,r] for j in I_total if j !=  'MC') == 1)
        # exactly one incoming arc to depot end node
        model.addConstr(quicksum(X[j,'MCd',r] for j in I_total if j != 'MCd') == 1)
        for i in I_total:
            for j in I_total:
                if (i != j) & ((i,j) != ('MC', 'MCd')):
                    # allow for routes not to be used by travelling from MC to MCd
                    # in that case, no other edge can be used on that route
                    # necessary? Doesn't SEC already prevent subtours? and there is no other possibility than a subtour to have the solver cheat, right?
                    model.addConstr(1 - X['MC', 'MCd',r] >= X[i,j,r])

    for i in I_total:
        if (i != 'MC') & (i != 'MCd'):
            # every node i has to be visited exactly once by the sum of all routes
            model.addConstr(quicksum(X[i,j,r] for r in range(allowed_routes) for j in I_total if j != i) == 1)  
            for r in range(allowed_routes):
                # every route r has to leave a node i as many times as it enters the node i
                model.addConstr(quicksum(X[i,j,r] for j in I_total if j != i) == quicksum(X[j,i,r] for j in I_total if j != i)) 


    """Time windows, Pick up and Drop off of the staff at clients (Home Health Care Problem)"""
    for i in (I1 + I2 + I3):
        model.addConstr(quicksum(W[i,s] for s in S if (i,s) in node_staff_tuple) == 1)
    
    for (i,s) in node_staff_tuple:
        if i != 'MCd':
            # only allow a Dt to be set if the staff is selected to work on client:
            model.addConstr(Dt[i,s] <= M * W[i,s]) # can this be replaced by D?
            model.addConstr(Pt[I_dict.get(i),s] <= M * W[i,s])
            # if the respective staff s is serving client i, it has to be dropped off by some route r:
            model.addConstr(quicksum(D[i,s,r] for r in range(allowed_routes)) == W[i,s]) 
            # if the respective staff s has been dropped off at client i, it needs to be picked up by some route r from the partner node ip:
            model.addConstr(quicksum(P[I_dict.get(i),s,r] for r in range(allowed_routes)) == W[i,s])
            for r in range(allowed_routes):
                # the drop off time of the respective staff s has to be equal to the arrival time of the route r at that respective drop off node i:
                model.addConstr((1 - D[i,s,r]) * M + Dt[i,s] >= t[i,r]) 
                model.addConstr(Dt[i,s] <= (1 - D[i,s,r]) * M + t[i,r])
                # the pick up time of the respective staff s has to be equal to the arrival time of the picking up route r at the respective node:
                model.addConstr(Pt[I_dict.get(i),s] <= (1 - P[I_dict.get(i),s,r]) * M + t[I_dict.get(i),r]) 
                model.addConstr((1 - P[I_dict.get(i),s,r]) * M + Pt[I_dict.get(i),s] >= t[I_dict.get(i),r])

    for i in I_total:
        for j in I_total:
            if i != j:
                for r in range(allowed_routes):
                    # if route r goes from i to j, the arrival time at j has to be larger or equal to the arrival time at i plus the travel time inbetween: MTZ-subtour elimination
                    model.addConstr((1 - X[i,j,r]) * M + t[j,r] >= t[i,r] + int(tt.get(i).get(j))) 
                    # kann hier künstlich Wartezeit erzeugt, aber nicht abgerechnet werden?
    if allow_overtime:
        # limit working hours of nurses to eight hours, additional time is overtime:
        model.addConstrs(Dt['MCd',s] - Pt['MC',s] <= 480 + overtime[s] for s in S)  
    else:
        # limit working hours of nurses to eight hours:
        model.addConstrs(Dt['MCd',s] - Pt['MC',s] <= 480 for s in S)  

    for s in S: 
        # limit working hours of nurses to ten hours:
        model.addConstr(Dt['MCd',s] - Pt['MC',s] <= 600)
        # every staff s can be at most on one route r and if so s needs to be picked up at MC:
        model.addConstr(quicksum(P['MC',s,r] for r in range(allowed_routes)) <= 1) 
        # if staff s is picked up, s also needs to be dropped off at MCd:
        model.addConstr(quicksum(P['MC',s,r] for r in range(allowed_routes)) == quicksum(D['MCd',s,r] for r in range(allowed_routes)))
        # staff s is always first picked up from MC before it can be dropped at MCd in the end:
        model.addConstr(Pt['MC',s] <= Dt['MCd',s])
        for r in range(allowed_routes):
            # set the drop off time at MC_ (equivalent to end of work):
            model.addConstr((1 - D['MCd',s,r]) * M + Dt['MCd',s] >= t['MCd',r]) 
            # set the pick up time at MC (equivalent to start of work):
            model.addConstr(Pt['MC',s] <= t['MC',r] + (1 - P['MC',s,r]) * M) 

    if allow_delay_clients:
        # set the delay of staff s at the client i with respect to the latest allowed starting time:
        model.addConstrs((1 - W[i,s]) * M + int(LST_client.get(i)) >= Dt[i,s] - delay_at_client[i] for (i,s) in node_staff_tuple if i != 'MCd')
    else:
        # ensure that the latest allowed starting time is always respected:
        model.addConstrs((1 - W[i,s]) * M + int(LST_client.get(i)) >= Dt[i,s] for (i,s) in node_staff_tuple if i != 'MCd')
    if allow_wait_staff:
        # keep track of how long staff s has to wait a client i before s can begin the service:
        model.addConstrs((W[i,s] - 1) * M + int(EST_client.get(i)) <= Dt[i,s] + wait_at_client_before[i,s] for (i,s) in node_staff_tuple if i != 'MCd')
        # difference between drop off time at client node i and pick up time at client partner node ip equals
        # the service duration plus the time s had to wait before and after the service:
        model.addConstrs((1 - W[i,s]) * M + Pt[I_dict.get(i),s] - Dt[i,s] == int(STD_client.get(i)) + wait_at_client_before[i,s] + wait_at_client_after[i,s] for (i,s) in node_staff_tuple if i != 'MCd')
        # every nurse can wait at most one hour a day:
        model.addConstrs(quicksum(wait_at_client_before[i,s] + wait_at_client_after[i,s] for i in I_total if (i,s) in node_staff_tuple) <= max_wait_staff for s in S) 
    else:
        # ensure that the earliest allowed starting time is always respected:
        model.addConstrs((W[i,s] - 1) * M + int(EST_client.get(i)) <= Dt[i,s] for (i,s) in node_staff_tuple if i != 'MCd')
        # difference between drop off time at client node i and pick up time at client partner node ip has to be
        # larger or equal to the standard service time:
        model.addConstrs((1 - W[i,s]) * M + Pt[I_dict.get(i),s] -  Dt[i,s] >= int(STD_client.get(i)) for (i,s) in node_staff_tuple if i != 'MCd')

    for (i,s) in node_staff_tuple:
        # if staff is dropped off at node i it has to be picked up at the respective pickup node ip:
        model.addConstr(quicksum(D[i,s,r] for r in range(allowed_routes)) == quicksum(P[I_dict.get(i),s,r] for r in range(allowed_routes))) 
        # only if the staff started at MC it can be sent to other nodes:
        model.addConstr(quicksum(P['MC',s,r] for r in range(allowed_routes)) >= quicksum(D[i,s,r] for r in range(allowed_routes)))
        # drop off time at MCd must be larger or equal than any pickup time of the staff (just making sure, the solver is not cheating):
        model.addConstr(Dt['MCd',s] >= Pt[I_dict.get(i),s])
        # similarly the pickup time at MC must be smaller or equal to any drop off time of the staff
        model.addConstr(Pt['MC',s] <= Dt[i,s])
       
    
    # Force only one time variable per node to take a value larger 0: NECCESSARY anymore???
    for i in I_total:
        if (i != 'MC') & (i != 'MCd'):
            model.addConstr(quicksum(y[i,r] for r in range(allowed_routes)) == 1) # jeder Knoten genau einmal angefahren
            for r in range(allowed_routes):
                model.addConstr(t[i,r] <= M * y[i,r])
                # model.addConstr(t[i,r] <= M * quicksum(X[i,j,r] for j in I_total if i!=j))
           
    for r in range(allowed_routes):
        for s in S:
            # every staff s that has been picked up by a route r has to be dropped off by the same route r
            model.addConstr(
                quicksum(P[I_dict.get(i),s,r] for i in I_total if (i,s) in node_staff_tuple) 
                == quicksum(D[i,s,r] for i in I_total if (i,s) in node_staff_tuple)
                )


    """Dial a ride section"""
    for i in I0:
        # every patient i needs to be picked up from home by a route r:
        model.addConstr(quicksum(p[i,r] for r in range(allowed_routes)) == 1) 
        for r in range(allowed_routes):
            # patient i can only be picked up on a route r that is going by his house:
            model.addConstr(p[i,r] <= quicksum(X[i,j,r] for j in I_total if i != j)) 
    for i in I_0: 
        # every patient i needs to be dropped off at home by a route r:
        model.addConstr(quicksum(d[i,r] for r in range(allowed_routes)) == 1) 
        for r in range(allowed_routes):
            # patient i can only be dropped off by a route r that is going by his house:
            model.addConstr(d[i,r] <= quicksum(X[i,j,r] for j in I_total if i != j)) 
    for r in range(allowed_routes):
        for i in I0:
            # if r is the picking up patient, the drop off time at MC corresponds to the arrival time of r at node MCd:
            model.addConstr(dt[i] >= t['MCd',r] - (1 - p[i,r]) * M) 
            model.addConstr(dt[i] <= t['MCd',r] + (1 - p[i,r]) * M) 
            # if r is dropping off patient, the pick up time at MC corresponds to the leaving time of r at node MC:
            model.addConstr(pt[I0_dict.get(i)] <= t['MC',r] + (1 - d[I0_dict.get(i),r]) * M) 
            model.addConstr(pt[I0_dict.get(i)] >= t['MC',r] - (1 - d[I0_dict.get(i),r]) * M)
    if allow_wait_patients:
        for i in I0:
            # respect time that it takes to process patient at MC:
            model.addConstr(int(STD_patient.get(i)) + wait_at_MC_after[i] + wait_at_MC_before[i] == pt[I0_dict.get(i)] - dt[i])
            # respect earliest starting time at MC (including waiting):
            model.addConstr(int(EST_patient.get(i)) <= dt[i] + wait_at_MC_before[i])
            # limit waiting time of patient i at MC to an hour:
            model.addConstr(wait_at_MC_before[i] + wait_at_MC_after[i] <= max_wait_patiens)
    else:
            for i in I0:
                # alternative without punishing wait time at MC:
                model.addConstr(int(STD_patient.get(i)) <= pt[I0_dict.get(i)] - dt[i])
                # respect earliest starting time at MC (without waiting):
                model.addConstr(int(EST_patient.get(i)) <= dt[i]) 
    for i in I0:
        # respect latest starting time at MC:
        model.addConstr(int(LST_patient.get(i)) >= dt[i])
        
    for i in I_total:
        if (i != 'MC') & (i != 'MCd'):
            for l in I0:
                for r in range(allowed_routes):
                    if i == l:
                        # at node i patient i must be picked up:
                        model.addConstr(quicksum(op[i,j,l,r] for j in I_total if i != j) == p[l,r]) # take this section down
                    elif i == I0_dict.get(l):
                        # at node i_ patient i must be dropped off:
                        model.addConstr(quicksum(op[i,j,l,r] for j in I_total if i != j) == 0) 
                    else:
                        # conserve flow of patient i:
                        model.addConstr(quicksum(op[i,j,l,r] for j in I_total if j != i) == quicksum(op[j,i,l,r] for j in I_total if i != j)) 

    """Vehicle Capacity Constraint"""
    for s in S:
        for r in range(allowed_routes):
            for i in I_total:
                model.addConstr(P['MC',s,r] == quicksum(o['MC',i,s,r] for i in I_total if i != 'MC'))
                if i != 'MC':
                    model.addConstr(o['MC',i,s,r] <= X['MC',i,r])
                if (i != 'MC') & (i != 'MCd'):
                    if (i,s) in node_staff_tuple:
                        # if staff s is on route r when travelling from i to j if it was on route r when arriving at i or it was picked up at node i:
                        model.addConstr(
                            quicksum(o[I_dict.get(i),j,s,r] for j in I_total if j != I_dict.get(i)) 
                            == quicksum(o[l,I_dict.get(i),s,r] for l in I_total if (l != I_dict.get(i)) & (l != 'MCd')) + P[I_dict.get(i),s,r]
                            )
                        # staff s is not in on route r if it was dropped off at node i:
                        model.addConstr(
                            quicksum(o[i,j,s,r] for j in I_total if j != i) 
                            == quicksum(o[l,i,s,r] for l in I_total if (l != i) & (l != 'MCd')) - D[i,s,r]
                            )
                    # if the respective node is neither a pick up nor a drop off node, the occupation o stays the same:
                    elif (find_key(i, I_dict),s) not in node_staff_tuple: 
                        model.addConstr(quicksum(o[i,j,s,r] for j in I_total if j != i) == quicksum(o[l,i,s,r] for l in I_total if (l != i) & (l != 'MCd'))) 
                    for j in I_total:
                        if i != j:
                            # only allow for staff s to be in traveling from i to j on route r if arc from i to j is used on route r:
                            model.addConstr(o[i,j,s,r] <= X[i,j,r])
    
    # if a staff s is transproted on a route r anywhere along the route, it needs to have been picked up at the MC by some route r:
    for s in S:
        for i in I_total:
            if (i != 'MC') & (i != 'MCd'):
                model.addConstr(
                    quicksum(o['MC',j,s,r] for j in I_total for r in range(allowed_routes) if j != 'MC') 
                    >= quicksum(o[i,j,s,r] for j in I_total for r in range(allowed_routes) if i != j)
                    )

    for l in I0: 
        for r in range(allowed_routes): 
            # force op to display the route on which the patient is dropped off:
            model.addConstr(quicksum(op['MC',j,l,r] for j in I_total if ('MC' != j) & ('MCd' != j)) == d[I0_dict.get(l),r]) 
            for i in I_total:
                for j in I_total:
                    if i != j:
                        # op can only be 1 if the route is either picking the patient up or dropping him off:
                        model.addConstr(op[i,j,l,r] <= p[l,r] + d[I0_dict.get(l),r]) 
                        # you can only transport a patient if the arc is used:
                        model.addConstr(op[i,j,l,r] <= X[i,j,r]) 
    
    for i in I_total:
        if i != 'MCd':
            for j in I_total:
                if i != j:
                    for r in range(allowed_routes):
                        # ensure that limited vehicle capacity is respected:
                        model.addConstr(quicksum(o[i,j,s,r] for s in S) + quicksum(op[i,j,l,r] for l in I0) <= vehicle_capacity)

    """Counting vehicles in parallel use"""  
    ## increases computational complexity by factor > 6
    # have increasing number of routes be equal to increasing time of leaving the depot: --> avoid symmetry (reduced computation time by roughly factor 5)
    for r1 in range(allowed_routes):
        for r2 in range(r1 + 1, allowed_routes):
            model.addConstr(predecessor[r1,r2] == 1)
            model.addConstr(predecessor[r2,r1] == 0)
    # if route r1 starts before route r2, and route r1 and route r2 are both performed on the same car c then route r1 has to end before route r2 starts:
    for r1 in range(allowed_routes):
        for r2 in range(allowed_routes):
            if r1 != r2:
                for c in range(number_of_vehicles):
                    model.addConstr(t['MCd', r1] <= t['MC', r2] + M * (3 - route_car_match[c,r1] - route_car_match[c,r2] - predecessor[r1,r2]))
    # only if a car c is being used, route r can be performed by it:
    for c in range(number_of_vehicles):
        for r in range(allowed_routes):
            model.addConstr(use_car[c] >= route_car_match[c,r]) 
    # avoid symmetry:
    for c1 in range(number_of_vehicles):
        for c2 in range(number_of_vehicles):
            if c2 > c1:
                model.addConstr(use_car[c1] >= use_car[c2]) 
    # every route has to be served by a car:
    for r in range(allowed_routes):
        model.addConstr(quicksum(route_car_match[c,r] for c in range(number_of_vehicles)) == 1)
    # # link binary car use variable to the number of cars being used:
    # model.addConstr(quicksum(use_car[c] for c in range(number_of_vehicles)) == cars_used)
    # set start und end time for specific car, as well es limit to one driver driving time
    for c in range(number_of_vehicles):
        model.addConstr(car_travel[c] == end_car_travel[c] - start_car_travel[c])
        model.addConstr(car_travel[c] <= 600)
        for r in range(allowed_routes):
            model.addConstr(start_car_travel[c] <= (1 - route_car_match[c,r]) * M + t['MC',r])
            model.addConstr((1 - route_car_match[c,r]) * M + end_car_travel[c] >= t['MCd',r])
            
    

    
    """Counting staff members used"""
    for (i,s) in node_staff_tuple:
        model.addConstr(use_staff[s] >= W[i,s])
    
    
    """---------------------------------------------------------------------------------------------------------------------"""
    """---------------------------------------------------------------------------------------------------------------------"""
    """---------------------------------------------------------------------------------------------------------------------"""
    """Optimization"""
    if allow_delay_clients and allow_overtime:
        model.setObjective(
            quicksum(use_car[c] for c in range(number_of_vehicles)) * (driver + car_fixed)  # fixed cost for car being used plus driver 8h shift
            + quicksum(use_staff[s] for s in S1) * level_1 
            + quicksum(use_staff[s] for s in S2) * level_2 
            + quicksum(use_staff[s] for s in S3) * level_3                                  # cost for nurses 8h shifts
            + quicksum(t['MCd',r] - t['MC',r] for r in range(allowed_routes)) * fuel_cost   # cost for fuel
            + quicksum(overtime[s] for s in S1) * (level_1/8 + over_time)/60                # cost for overtime (per minte)
            + quicksum(overtime[s] for s in S2) * (level_2/8 + over_time)/60 
            + quicksum(overtime[s] for s in S3) * (level_3/8 + over_time)/60
            + quicksum(delay_at_client[i] for i in (I1 + I2 + I3)) *20/60                   # penalty cost for being late
            )
    elif (not allow_delay_clients) and allow_overtime:
        model.setObjective(
            quicksum(use_car[c] for c in range(number_of_vehicles)) * (driver + car_fixed)  # fixed cost for car being used plus driver 8h shift
            + quicksum(use_staff[s] for s in S1) * level_1 
            + quicksum(use_staff[s] for s in S2) * level_2 
            + quicksum(use_staff[s] for s in S3) * level_3                                  # cost for nurses 8h shifts
            + quicksum(t['MCd',r] - t['MC',r] for r in range(allowed_routes)) * fuel_cost   # cost for fuel
            + quicksum(overtime[s] for s in S1) * (level_1/8 + over_time)/60                # cost for overtime (per minte)
            + quicksum(overtime[s] for s in S2) * (level_2/8 + over_time)/60 
            + quicksum(overtime[s] for s in S3) * (level_3/8 + over_time)/60
            )
    elif allow_delay_clients and (not allow_overtime):
        model.setObjective(
            quicksum(use_car[c] for c in range(number_of_vehicles)) * (driver + car_fixed)  # fixed cost for car being used plus driver 8h shift
            + quicksum(use_staff[s] for s in S1) * level_1 
            + quicksum(use_staff[s] for s in S2) * level_2 
            + quicksum(use_staff[s] for s in S3) * level_3                                  # cost for nurses 8h shifts
            + quicksum(t['MCd',r] - t['MC',r] for r in range(allowed_routes)) * fuel_cost   # cost for fuel
            + quicksum(delay_at_client[i] for i in (I1 + I2 + I3)) *20/60                   # penalty cost for being late
            )
    else:
        model.setObjective(
            quicksum(use_car[c] for c in range(number_of_vehicles)) * (driver + car_fixed)  # fixed cost for car being used plus driver 8h shift
            + quicksum(use_staff[s] for s in S1) * level_1 
            + quicksum(use_staff[s] for s in S2) * level_2 
            + quicksum(use_staff[s] for s in S3) * level_3                                  # cost for nurses 8h shifts
            + quicksum(t['MCd',r] - t['MC',r] for r in range(allowed_routes)) * fuel_cost   # cost for fuel
            )
    # log_callback = LogCallback()
    model.setParam(GRB.Param.TimeLimit, 1800)
    # model.setParam(GRB.Param.MIPFocus, 2)
    model.setParam(GRB.Param.SolutionLimit, 25)
    model.update()
    # model.optimize(LogCallback.__call__(log_callback, model, GRB.Callback.MIP))
    model.optimize()

    """Output"""
    if model.Status == GRB.OPTIMAL:
    
        """Simple Pre-Checks"""
        for i in I_total:
            for j in I_total:
                for s in S:
                    for r in range(allowed_routes):
                        if (i != j) & (i != 'MCd'):
                            # check wheter staff is traveling on arcs that are not used in a route:
                            if round(o[i,j,s,r].getAttr('X')) > round(X[i,j,r].getAttr('X')):
                                print(f'Staff {s} is travelling on route {r} form {i} to {j}, although the arc is not used in that route!')
        
        for i in I_total:
            for s in S:
                for r in range(allowed_routes):
                    # check if flow conversation is met:
                    outgoing_sum = sum(o[i,l,s,r].getAttr('X') for l in I_total if (i != l) & (i != 'MCd'))
                    incoming_sum = sum(o[l,i,s,r].getAttr('X') for l in I_total if (i != l) & (l != 'MCd'))
                    if round(outgoing_sum) > round(incoming_sum):
                        if round(P[i,s,r].getAttr('X')) < 1:
                            print(f'Staff {s} is picked up at node {i} on route {r}, although the pick up variable is not set correspondently.')
                    if round(incoming_sum) > round(outgoing_sum):   
                        if round(D[i,s,r].getAttr('X')) < 1:
                            print(f'Staff {s} is dropped off at node {i} on route {r}, although the drop off variable is not set correspondently.')                     

        """Export Staff Schedule"""
        for s in S:
            data = {'Current Node': [],
                  'Next Node': [],
                  'Time': [],
                  'Drop off?': [],
                  'EST': [],
                  'LST': [],
                  'Task duration': [],
                  'Pick up?': [],
                  'Route:': [],
                  }
            file_name = f'outputs/schedule_staff_{s}.xlsx'
            df = pd.DataFrame(data)
            output_file_path = f'schedule_{s}'
            staff_used = False
            for i in I_total:
                for j in I_total:
                    if (i!=j) & (i!='MCd'):
                        for r in range(allowed_routes):
                            if round(o[i,j,s,r].getAttr('X')) == 1:
                                if j != 'MCd':
                                    staff_used = True
                                if (i,s) in node_staff_tuple:
                                    if round(D[i,s,r].getAttr('X')) == 1:
                                        new_row =   {'Current Node': i,
                                                    'Next Node': j,
                                                    'Time': round(t[i,r].getAttr('X')),
                                                    'Drop off?': 'yes',
                                                    'EST': EST_client.get(i),
                                                    'LST': LST_client.get(i),
                                                    'Task duration': STD_client.get(i),
                                                    'Pick up?': 'no',
                                                    'Route:': 'working',
                                                    }
                                    else:
                                        new_row =   {'Current Node': i,
                                                    'Next Node': j,
                                                    'Time': round(t[i,r].getAttr('X')),
                                                    'Drop off?': 'no',
                                                    'EST': 'n.a.',
                                                    'LST': 'n.a.',
                                                    'Task duration': 'n.a.',
                                                    'Pick up?': 'no',
                                                    'Route:': r,
                                                    }
                                if ((j,s) in node_staff_tuple) & (j!='MCd'):
                                    if round(D[j,s,r].getAttr('X')) == 1:
                                        new_row_2 = {'Current Node': j,
                                                    'Next Node': I_dict.get(j),
                                                    'Time': round(t[j,r].getAttr('X')),
                                                    'Drop off?': 'yes',
                                                    'EST': EST_client.get(j),
                                                    'LST': LST_client.get(j),
                                                    'Task duration': STD_client.get(j),
                                                    'Pick up?': 'no',
                                                    'Route:': 'working',
                                                    }
                                        connect_df_2 = pd.DataFrame([new_row_2])
                                        df = pd.concat([df, connect_df_2])
                                        new_row =   {'Current Node': i,
                                                    'Next Node': j,
                                                    'Time': round(t[i,r].getAttr('X')),
                                                    'Drop off?': 'no',
                                                    'EST': 'n.a.',
                                                    'LST': 'n.a.',
                                                    'Task duration': 'n.a.',
                                                    'Pick up?': 'no',
                                                    'Route:': r,
                                                    }
                                if (find_key(i, I_dict),s) in node_staff_tuple:
                                    if round(P[i,s,r].getAttr('X')) == 1:
                                        new_row =   {'Current Node': i,
                                                    'Next Node': j,
                                                    'Time': round(t[i,r].getAttr('X')),
                                                    'Drop off?': 'no',
                                                    'EST': 'n.a.',
                                                    'LST': 'n.a.',
                                                    'Task duration': 'n.a.',
                                                    'Pick up?': 'yes',
                                                    'Route:': r,
                                                    }
                                    else:
                                        new_row =   {'Current Node': i,
                                                    'Next Node': j,
                                                    'Time': round(t[i,r].getAttr('X')),
                                                    'Drop off?': 'no',
                                                    'EST': 'n.a.',
                                                    'LST': 'n.a.',
                                                    'Task duration': 'n.a.',
                                                    'Pick up?': 'no',
                                                    'Route:': r,
                                                    }
                                else:
                                    new_row =   {'Current Node': i,
                                                'Next Node': j,
                                                'Time': round(t[i,r].getAttr('X')),
                                                'Drop off?': 'no',
                                                'EST': 'n.a.',
                                                'LST': 'n.a.',
                                                'Task duration': 'n.a.',
                                                'Pick up?': 'no',
                                                'Route:': r,
                                                }
                                if j == 'MCd':
                                    new_row_2 =   {'Current Node': j,
                                    'Next Node': 'n.a.',
                                    'Time': round(t[j,r].getAttr('X')),
                                    'Drop off?': 'yes',
                                    'EST': 'n.a.',
                                    'LST': 'n.a.',
                                    'Task duration': 'n.a.',
                                    'Pick up?': 'no',
                                    'Route:': r,
                                    }
                                    connect_df_2 = pd.DataFrame([new_row_2])
                                    df = pd.concat([df, connect_df_2])
                                connect_df = pd.DataFrame([new_row])
                                df = pd.concat([df, connect_df])
            # sort entries by time
            df_sorted = df.sort_values(by='Time')
            df_sorted.set_index(['Time'], inplace=True)
            if staff_used:
                # save output to excel file
                df_sorted.to_excel(file_name, engine='openpyxl')
                workbook = oxl.load_workbook(file_name)
                worksheet = workbook.active

                # adjust column witdh and center entries
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter  # Get the column name
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    adjusted_width = (max_length + 2)  # Extra space for better readability
                    worksheet.column_dimensions[column].width = adjusted_width
                workbook.save(file_name)
                
        
        """Derive Car Schedule"""
        for c in range(number_of_vehicles):
            if round(use_car[c].getAttr('X')) == 0:
                pass
            else: 
                file_name = f'outputs/schedule_car_{c}.xlsx'
                data = {'Current Node': [],
                        'Next Node': [],
                        'Time': [],
                        'Used Capacity (leaving the node)': [],
                        'Route:': [],
                        }
                df = pd.DataFrame(data)
                for r in range(allowed_routes):
                    if round(X['MC','MCd',r].getAttr('X')) == 0:
                        if round(route_car_match[c,r].getAttr('X')) == 1:
                            new_row =   {'Current Node': 'MCd',
                                        'Next Node': 'n.a.',
                                        'Time': round(t['MCd',r].getAttr('X')),
                                        'Used Capacity (leaving the node)': 'n.a.',
                                        'Route:': r,
                                        }
                            connect_df = pd.DataFrame([new_row])
                            df = pd.concat([df, connect_df])
                            for i in I_total: 
                                for j in I_total:
                                    if i != j:
                                        if round(X[i,j,r].getAttr('X')) == 1:
                                            new_row =   {'Current Node': i,
                                                        'Next Node': j,
                                                        'Time': round(t[i,r].getAttr('X')),
                                                        'Used Capacity (leaving the node)': sum(round(o[i,j,s,r].getAttr('X')) for s in S) 
                                                                                            + sum(round(op[i,j,l,r].getAttr('X')) for l in I0),
                                                        'Route:': r,
                                                        }
                                            connect_df = pd.DataFrame([new_row])
                                            df = pd.concat([df, connect_df])
                # sort entries by time 
                df_sorted = df.sort_values(by='Time')
                df_sorted.set_index(['Time'], inplace=True)
                # save output to excel file
                df_sorted.to_excel(file_name, engine='openpyxl')
                workbook = oxl.load_workbook(file_name)
                worksheet = workbook.active

                # adjust column witdh and center entries
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter  # Get the column name
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    adjusted_width = (max_length + 2)  # Extra space for better readability
                    worksheet.column_dimensions[column].width = adjusted_width
                workbook.save(file_name)

                        

        variable_groups = {}
        for var in model.getVars():
            name = var.varName.split('_')[0]  # Extract variable name without index
            if name not in variable_groups:
                variable_groups[name] = []
            variable_groups[name].append(var)

        # Save output to a file
        output_file_path = "outputs/decision_variables.txt"
        nonzero_output_file_path = "outputs/nonzero_decision_variables.txt"
        with open(output_file_path, "w") as output_file:

            # Print decision variables grouped by their names to the file
            for name, vars in variable_groups.items():
                output_file.write(f"Variables with name '{name}':\n")
                for var in vars:
                    output_file.write(f"{var.varName} = {var.x}\n")

        with open(nonzero_output_file_path, "w") as output_file:

            # Print decision variables grouped by their names to the file
            for name, vars in variable_groups.items():
                output_file.write(f"Variables with name '{name}':\n")
                for var in vars:
                    if var.x >= 0.99:
                        output_file.write(f"{var.varName} = {np.round(var.x,0)}\n")

        arcs = []
        for i in I_total:
            for j in I_total:
                if i != j:
                    for r in range(allowed_routes):
                        if X[i,j,r].getAttr('X') >= 0.99:
                            arcs += [(i,j,r)]
        node_info = {'MC': '',
                     'MCd': '',
                     }
        for i in I_total:
            if (i != 'MC') & (i != 'MCd'):
                for s in S:
                    if (i,s) in node_staff_tuple:
                        if Dt[i,s].getAttr('X') > 0.01:
                            node_info[i] = np.round(Dt[i,s].getAttr('X'), 0)
                        if Pt[I_dict.get(i),s].getAttr('X') > 0.01:
                            node_info[I_dict.get(i)] = np.round(Pt[I_dict.get(i),s].getAttr('X'), 0)

        return model.Status, arcs, node_info
    elif model.Status == GRB.TIME_LIMIT:
        print(f"Optimization stopped, because Time Limit is reached. Optimilaity Gap: {model.MIPGap}")
        return model.Status, None, None
    elif model.Status == GRB.INFEASIBLE:
        print("Model is infeasible.")
        model.computeIIS()
        model.write("infeasibility_report.ilp")
        return model.Status, None, None
    



"""TODOs"""
    # Do we have symmetry?
    # Working hours drivers --> included
    # limit waiting time of patient --> included
    # break time of nurses
    # limit working time and unproductive time of nurses --> included
    # what is a good big M?
    # avoid staff being on two tours at the same time --> included
    # how many drivers and cars are used at the same time? Which consecutive tours could be performed by the same car, driver, nurse --> included for cars and drivers
    # retrieve the final plan for the operator
    # allow for selection and deselection of model flexibilities
    # allow for clients not to be served at all?
    # exclude edges from i to ip for faster solving?
    # compare MTZ to DFJ?


"""Create Graphs showing Nodes investigated, Open Nodes as well as Primal and Dual Bound over optimization time"""
# class LogCallback:
#     def __init__(self):
#         self.primal_bounds = []
#         self.dual_bounds = []
#         self.open_nodes = []
#         self.processed_nodes = []

#     def __call__(self, model, where):
#         if where == GRB.Callback.MIP:
#             # Get primal and dual bounds
#             primal_bound = model.cbGet(GRB.Callback.MIP_OBJBST)
#             dual_bound = model.cbGet(GRB.Callback.MIP_OBJBND)
#             # Get number of open and processed nodes
#             open_nodes = model.cbGet(GRB.Callback.MIP_NODLFT)
#             processed_nodes = model.cbGet(GRB.Callback.MIP_NODCNT)

#             # Store values
#             self.primal_bounds.append(primal_bound)
#             self.dual_bounds.append(dual_bound)
#             self.open_nodes.append(open_nodes)
#             self.processed_nodes.append(processed_nodes)


"""Tiny function to find the key of an dictionary entry"""
def find_key(value_to_find, dict):
    for key, value in dict.items():
        if value == value_to_find:
            return key
    return None