
from openpyxl import load_workbook

load_wb = load_workbook("calc_error.xlsx")
load_ws = load_wb['Sheet1']
best_n = load_ws['B3': 'B35']
best_s = load_ws['D3': 'D35']
bests = {}
for index in range(len(best_n)):
    name = best_n[index][0].value
    score = best_s[index][0].value
    bests[name] = score
    # print(name, score)

run_n = load_ws['F3': 'F40']
run_s = load_ws['H3': 'H40']
runs = {}
for index in range(len(run_n)):
    name = run_n[index][0].value
    score = run_s[index][0].value
    runs[name] = score
    # print(name, score)

inf_n = load_ws['J3': 'J37']
inf_s = load_ws['L3': 'L37']
infs = {}
for index in range(len(inf_n)):
    name = inf_n[index][0].value
    score = inf_s[index][0].value
    infs[name] = score
    # print(name, score)

sum = 0
######
#best - run
for key in bests.keys():
    left = bests[key]
    right = 0
    if key in runs:
        right = runs[key]

    # print(key, right)

    sum += abs(left - right)

result = sum / len(best_n)
print("best-run MAE is " + str(result))


######
#best - infs

sum = 0
######
#best - run
for key in bests.keys():
    left = bests[key]
    right = 0
    if key in infs:
        right = infs[key]
    sum += abs(left - right)

result = sum / len(best_n)
print("best-inf MAE is " + str(result))

print(abs(-1), abs(1))
