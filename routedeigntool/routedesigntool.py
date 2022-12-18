import xlwings as xw
from xlwings import Range, constants
import openpyxl as op
"""Simple Travelling Salesperson Problem (TSP) between cities."""
import pandas as pd
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp

wb = xw.Book('routedesigntool.xls')



#tao diem dau tien co gia tri trung voi ten sheet

ws1 = wb.sheets[1]

#wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row

nhap1 = int(ws1.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').value)

print(nhap1)

def create_data_model():
    """Stores the data for the problem."""
    data = {}
    data['distance_matrix'] = [
        [0, 50, 100, 150, 210, 225, 140, 130, 360, 210, 210, 280, 310, 350, 390, 430, 480, 500, 510, 530, 230, 210, 150], 
[62, 0, 50, 70, 130, 200, 130, 110, 310, 160, 160, 230, 260, 300, 340, 380, 430, 450, 460, 480, 180, 160, 100], 
[100, 50, 0, 110, 100, 150, 60, 50, 260, 110, 110, 180, 210, 250, 290, 330, 380, 400, 410, 430, 130, 110, 250], 
[160, 80, 70, 0, 73, 70, 60, 110, 340, 190, 190, 150, 180, 220, 260, 300, 350, 370, 380, 400, 90, 100, 170], 
[210, 140, 100, 60, 0, 10, 90, 100, 270, 120, 120, 80, 110, 150, 190, 230, 280, 300, 310, 330, 130, 150, 210], 
[220, 150, 110, 70, 10, 0, 140, 110, 280, 130, 80, 70, 100, 140, 180, 220, 270, 290, 300, 320, 140, 160, 220], 
[140, 130, 60, 120, 80, 70, 0, 30, 230, 40, 70, 120, 150, 190, 230, 270, 320, 340, 350, 370, 200, 210, 270], 
[130, 120, 50, 110, 80, 70, 10, 0, 240, 50, 80, 130, 160, 200, 240, 280, 330, 350, 360, 380, 190, 200, 260], 
[300, 340, 140, 220, 160, 150, 130, 150, 0, 90, 130, 150, 180, 220, 260, 300, 350, 370, 380, 400, 300, 310, 370], 
[180, 118, 100, 110, 60, 50, 30, 40, 150, 0, 40, 130, 160, 200, 240, 280, 330, 350, 360, 380, 180, 190, 250], 
[210, 160, 110, 130, 70, 60, 70, 80, 240, 140, 0, 160, 190, 230, 270, 310, 360, 380, 390, 410, 200, 210, 270], 
[350, 410, 260, 220, 130, 120, 210, 200, 360, 220, 170, 0, 30, 70, 110, 150, 200, 220, 230, 250, 120, 130, 190], 
[330, 390, 240, 200, 110, 100, 190, 180, 340, 200, 150, 20, 0, 40, 80, 120, 170, 190, 200, 220, 100, 110, 170], 
[370, 420, 300, 270, 280, 290, 350, 330, 500, 380, 300, 320, 360, 0, 40, 80, 130, 150, 160, 180, 230, 240, 300], 
[330, 380, 260, 230, 240, 250, 310, 290, 460, 340, 260, 280, 320, 360, 0, 40, 90, 110, 120, 140, 190, 200, 260], 
[290, 340, 220, 190, 200, 210, 270, 250, 420, 300, 220, 240, 280, 320, 360, 0, 50, 70, 80, 100, 150, 160, 220], 
[240, 290, 170, 140, 150, 160, 220, 200, 370, 250, 170, 190, 230, 270, 310, 360, 0, 20, 30, 50, 100, 110, 170], 
[220, 270, 150, 120, 130, 140, 200, 180, 350, 230, 150, 170, 210, 250, 290, 340, 360, 0, 40, 60, 80, 90, 150], 
[210, 260, 140, 110, 120, 130, 190, 170, 340, 220, 140, 160, 200, 240, 280, 330, 350, 360, 0, 30, 40, 50, 110], 
[190, 240, 120, 90, 100, 110, 170, 150, 320, 200, 120, 140, 180, 220, 260, 310, 330, 340, 360, 0, 50, 60, 120], 
[140, 180, 150, 80, 130, 140, 190, 180, 420, 470, 550, 210, 250, 290, 330, 380, 400, 410, 430, 310, 0, 10, 70], 
[130, 170, 140, 70, 120, 130, 180, 170, 410, 460, 540, 200, 240, 280, 320, 370, 390, 400, 420, 300, 10, 0, 30], 
[100, 140, 110, 40, 90, 100, 150, 140, 380, 430, 510, 170, 210, 250, 290, 340, 360, 370, 390, 270, 20, 30, 0]
    ]  # yapf: disable
    data['num_vehicles'] = 1
    data['depot'] = nhap1
    
    return data
def print_solution(manager, routing, solution):
    """Prints solution on console."""
    print('Objective: {} seconds'.format(solution.ObjectiveValue()))
    index = routing.Start(0)
    plan_output = 'Route for vehicle 0:\n'
    plan_output1 = []
    route_distance = 0
    while not routing.IsEnd(index):
        plan_output += ' {} \n'.format(manager.IndexToNode(index))

        previous_index = index
        index = solution.Value(routing.NextVar(index))
        route_distance += routing.GetArcCostForVehicle(previous_index, index,0)
        thoigian = routing.GetArcCostForVehicle(previous_index, index,0)
        plan_output1.append(int(previous_index))
        
    print(plan_output1)
    plan_output += ' {}\n'.format(manager.IndexToNode(index))
    plan_output += 'Route distance: {} seconds\n'.format(route_distance)
    tong_time = route_distance - thoigian
    print(tong_time)



    
    #in ra excel
    nhap1 = int(ws1.range('A1').expand().last_cell.row)
    ws1.range('B'+str(nhap1)).value = tong_time
    ws1.range('C'+str(nhap1)).value = plan_output1    
  

def main():
    """Entry point of the program."""
    # Instantiate the data problem.
    data = create_data_model()

    # Create the routing index manager.
    manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                           data['num_vehicles'], data['depot'])
   
    # Create Routing Model.
    routing = pywrapcp.RoutingModel(manager)
   


    def distance_callback(from_index, to_index):
        """Returns the distance between the two nodes."""

        

        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return data['distance_matrix'][from_node][to_node]

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)

    # Define cost of each arc.
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    # Setting first solution heuristic.
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    #Giai thuat greedy
    search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.FIRST_UNBOUND_MIN_VALUE)

    #giai thuat 2
    #search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.LOCAL_CHEAPEST_ARC)

    #giai thuat 3
    #search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.PARALLEL_CHEAPEST_INSERTION)

    #giai thuat 4 == giai thuat 3
    #search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.LOCAL_CHEAPEST_INSERTION)

    #giai thuat 5 == giai thuat 4 == giai thuat 3
    #search_parameters.first_solution_strategy = (routing_enums_pb2.FirstSolutionStrategy.GLOBAL_CHEAPEST_ARC)

    #giai thuat 6

    # search_parameters.local_search_metaheuristic = (routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH)
    # search_parameters.time_limit.seconds = 30
    # search_parameters.log_search = True

     #giai thuat 7

    #search_parameters.local_search_metaheuristic = (routing_enums_pb2.LocalSearchMetaheuristic.SIMULATED_ANNEALING)
    #search_parameters.time_limit.seconds = 10
    #search_parameters.log_search = True


    # Solve the problem.
    solution = routing.SolveWithParameters(search_parameters)

    # Print solution on console.
    if solution:
        print_solution(manager, routing, solution)
if __name__ == '__main__':
    main()          
    
       





