# AGAP

# Test1 is the main script. It uses the optimization techniques from the OR-Tools N-Queens Problem. Data is read from the excel file "Schedule" and returned through the schedule_data() function. 

#Currently the programme prints a single solution for the gate assignment where the only constraint is that each assignment must be on a different row in the matrix. The other constraints are blocked out at the moment as they have indexing errors. Each of the constraints are described in the programme. These constraints must be implemented first, then the OR-Tools solver has in-built functions for the objective function - To minimise walking distance. 

#I will work on the constraints this weekend / early next week, and will then implement the objective function and check the limits for using the airport for stopover operations.
