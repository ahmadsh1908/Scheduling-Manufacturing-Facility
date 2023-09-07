import xlsxwriter
from ortools.linear_solver import pywraplp  # calling goolge or tools library

# First Quarter
SA = [1, 2, 3, 4]  # Assigning for SA indexes=1..4
JW = [5, 6, 7, 8]  # Assigning for JW indexes=5..8
SW = [9, 10, 11, 12]  # Assigning for SW indexes=9..12
workers = range(1, 13)  # worker range from 1--12 (13-1)
weeks = range(1, 53)  # 52 weeks per year
days = range(1, 8)  # 7 days per week
shift = range(1, 8)  # 7 shifts
weekdays = range(1, 6)  # 5 days during weekdays. 1=Monday... 5=Friday
weekends = range(6, 8)  # 2 days on weekends 6=Saturday, 7=Sunday

solver = pywraplp.Solver('simple_Mixed Integer_program',
                         pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)  # Calling CBC solver since Integer Linear Program.

shifts = {}
a = {}  # Balance variable for weekdays shift
b = {}  # Balance variable for weekend shift

for i in workers:
    for w in weeks:
        for d in days:
            for s in shift:
                # declaring our 4 dimensional boolean matrix, 1 if work on week w day d shift s, 0 Otherwise.
                shifts[(i, w, d, s)] = solver.BoolVar(
                    'shift_i%iw%id%is%i' % (i, w, d, s))

for i in range(1, 4):  # One balance variable for each worker category
    a[i] = solver.NumVar(0, solver.infinity(), 'a_i%i' % (i))
    b[i] = solver.NumVar(0, solver.infinity(), 'b_i%i' % (i))

for i in SA:
    # Balacing SA category
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekdays for s in shift) <= a[1])
for i in JW:
    # Balancing Junior Workers
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekdays for s in shift) <= a[2])
for i in SW:
    # Balancing Senior Workers
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekdays for s in shift) <= a[3])

for i in SA:
    # Balacing SA category
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekends for s in shift) <= b[1])
for i in JW:
    # Balancing Junior Workers
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekends for s in shift) <= b[2])
for i in SW:
    # Balancing Senior Workers
    solver.Add(sum(shifts[(i, w, d, s)]
               for w in weeks for d in weekends for s in shift) <= b[3])

# Shifts can only be done once per day
for w in weeks:
    for d in days:
        if d < 6:
            for s in shift:
                solver.Add(sum(shifts[(i, w, d, s)] for i in workers) == 1)
        else:
            for s in shift:
                if s < 3 or s > 6:
                    # Weekend only C1-C2-C7 are performed
                    solver.Add(sum(shifts[(i, w, d, s)] for i in workers) == 1)
                else:
                    # Weekend only Shifts C3-C4-C5-C6 cannot be done
                    solver.Add(sum(shifts[(i, w, d, s)] for i in workers) == 0)

# Each worker works at most one shift per day.
for i in workers:
    for w in weeks:
        for d in days:
            solver.Add(sum(shifts[(i, w, d, s)] for s in shift) <= 1)

# If C7 then nextday off
for i in workers:
    for w in weeks:
        for d in days:
            if d == 7:  # If sunday then Monday (1) off
                if w < len(weeks):  # if the current week is not the last
                    solver.Add(
                        shifts[(i, w, d, 7)] + sum(shifts[(i, w+1, 1, s)] for s in shift) <= 1)
            else:
                solver.Add(shifts[(i, w, d, 7)] +
                           sum(shifts[(i, w, d+1, s)] for s in shift) <= 1)


# Balancing Categories
for i in workers:
    if i <= 4:  # a. Senior Associate (SA): # SA Workers(1,2,3,4)
        for w in weeks:
            for d in days:
                # Weekdays
                if d < 6:
                    # SA Workers Only can work C1 and C4 During weekdays
                    solver.Add(shifts[(i, w, d, 1)] +
                               shifts[(i, w, d, 4)] <= 1)
                    # SA Workers Can't work shifts C2,C3,C5,C6,C7
                    solver.Add(shifts[(i, w, d, 2)] + shifts[(i, w, d, 3)] + shifts[(i, w, d, 5)] +
                               shifts[(i, w, d, 6)] + shifts[(i, w, d, 7)] == 0)
                # Weekends
                else:
                    # SA Workers Only can work C1 and C2 During weekends
                    solver.Add(shifts[(i, w, d, 1)] +
                               shifts[(i, w, d, 2)] <= 1)
                    # SA Workers can't work C3 - C7 During weekends
                    solver.Add(sum(shifts[(i, w, d, s)]
                               for s in range(3, 8)) == 0)
    elif i <= 8 and i > 4:  # b. Junior Workers (JW):  # JW Workers(5,6,7,8)
        for w in weeks:
            for d in days:
                # Weekdays
                if d < 6:
                    # JW Workers Only can work C2, C3, C5, C6, C7  During weekdays
                    solver.Add(shifts[(i, w, d, 2)] + shifts[(i, w, d, 3)] + shifts[(i, w, d, 5)]
                               + shifts[(i, w, d, 6)] + shifts[(i, w, d, 7)] <= 1)
                    # JW Workers Can't work shifts C1,C4
                    solver.Add(shifts[(i, w, d, 1)] +
                               shifts[(i, w, d, 4)] == 0)
                # Weekends
                else:
                    # JW Only can work  C1, C2, C7  During Weekends
                    solver.Add(
                        shifts[(i, w, d, 1)] + shifts[(i, w, d, 2)] + shifts[(i, w, d, 7)] <= 1)
                    # JW Workers can't work C3-C6 During Weekends
                    solver.Add(sum(shifts[(i, w, d, s)]
                               for s in range(3, 7)) == 0)

    elif i > 8 and i <= 12:
        for w in weeks:
            for d in days:
                # Weekdays
                if d < 6:
                    # SW Only can work C1, C2, C3, C4, C5, C6  During weekdays
                    solver.Add(sum(shifts[(i, w, d, s)]
                               for s in range(1, 7)) <= 1)
                    # SW Can't work shifts C7
                    solver.Add(shifts[(i, w, d, 7)] == 0)
                    # Weekends
                else:
                    # SW Only can work  C1, C2 During weekends
                    solver.Add(shifts[(i, w, d, 1)] +
                               shifts[(i, w, d, 2)] <= 1)
                    # SW can't work C3 - C7 During weekdays
                    solver.Add(sum(shifts[(i, w, d, s)]
                               for s in range(3, 8)) == 0)


# Objective Function is to minimize the balancing variable in order to get a balanced worked schedule.
solver.Minimize(sum(a[i] for i in range(1, 4))+sum(b[i] for i in range(1, 4)))

status = solver.Solve()


names = ["", "CR", "FM", "MY", "BB", "JB",
         "RK", "CK", "CY", "HT", "AY", "CA", "GK"]
Weekdays = ["", "Monday", "Tuesday", "Wednesday",
            "Thursday", "Friday", "Saturday", "Sunday"]
z = " "


workbook = xlsxwriter.Workbook('YearSolution_BalancedWithin.xlsx')

worksheet = workbook.add_worksheet("Year")

title = workbook.add_format()
title.set_align('center')
title.set_bold()

worksheet.write(0, 5, "Year 2021", title)

row = 2
col = 1

style = workbook.add_format()
style.set_border(True)
style.set_align('center')

dayStyle = workbook.add_format()
dayStyle.set_align('center')

worksheet.set_column(0, 0, 6)
worksheet.set_column(1, 1, 12)

for w in weeks:
    worksheet.write(row, col, "WEEK "+str(w), title)
    for s in shift:
        col += 1
        worksheet.write(row, col, "C"+str(s), title)
    for d in days:
        row += 1
        col = 1
        worksheet.write(row, col, str(Weekdays[d]), dayStyle)
        for s in shift:
            i = 1
            while i < 13:
                if (shifts[(i, w, d, s)].solution_value() >= 1):
                    z = str(names[i])
                    break
                i += 1
            col += 1
            worksheet.write(row, col, z, style)
            z = "--"
        col = 1
    row += 2
workbook.close()

# counting Shifts per Category
print("Total Shifts per Category")

print("Senior Associates Shifts=", end=" ")
print(sum(shifts[(i, w, d, s)].solution_value()
      for i in SA for w in weeks for d in days for s in shift))

print("Junior Workers Shifts=", end=" ")
print(sum(shifts[(i, w, d, s)].solution_value()
      for i in JW for w in weeks for d in days for s in shift))

print("Senior Workers Shifts=", end=" ")
print(sum(shifts[(i, w, d, s)].solution_value()
      for i in SW for w in weeks for d in days for s in shift))

print('\n')


# Counting Shifts per Worker on Weekdays

print("Total Shifts per Worker on Weekdays")

print("Senior Associates")
for i in SA:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekdays for s in shift))

print("Junior Workers")
for i in JW:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekdays for s in shift))

print("Senior Workers")

for i in SW:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekdays for s in shift))


# Counting Shifts per Worker on Weekends
print('\n')
print("Total Shifts per Worker on Weekends")

print("Senior Associates")
for i in SA:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekends for s in shift))

print("Junior Workers")
for i in JW:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekends for s in shift))

print("Senior Workers")

for i in SW:
    print('Total Shifts worked by', names[i], end="=")
    print(sum(shifts[(i, w, d, s)].solution_value()
          for w in weeks for d in weekends for s in shift))
