import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Step 1
# ----------------------------------------------------------------------------------------------------------------------
matrix = pd.read_excel('APU_5_2017_soybeans_clip (1).xlsx', header=5, engine='openpyxl')
matrix.to_excel('Linear Optimization - APU_5_2017_soybeans_clip (1).xlsx')

list1 = matrix.stack()  # Flattens Dataframe into Series
std = round(list1.std(), 4)
list2 = list1.sort_values()  # Sorts the values in ascending order

print(list1.to_string())
print("\n\n\n\n\n\n")
print(list2.to_string())
print("\n\n\n\n\n\n")

listapp = []
list3 = np.array_split(list2, 2)  # Splits the list of numbers into two lists
sum = 0.0
for i in list3:
    for j in i:
        sum = sum + j

print(list3)
print("\n\n\n\n\n\n")
list3_0 = list3[0]
list3_1 = list3[1]
#  ^this splits list3 into two arrays (this was necessary as I could not index a nested list)

sum2_0 = 0
sum2_0_1 = 0
for item in list3_0:
    y1 = item

    x = 1
    m1 = (y1 - 50) / x
    # print(m1)
    if m1 < 0:
        y1_0 = 50
    else:
        y1_0 = 50 + (m1 * 1)  # Fertilizer application not adjusted
    y1 = 50 + (m1 * 0)  # Fertilizer application adjusted
    print(y1)
    listapp.append(y1)
    sum2_0 = sum2_0 + y1  # Sums all of the results (after process)
    sum2_0_1 = sum2_0_1 + y1_0  # Sums all of the results (before process)

sum2_1 = 0
sum2_1_1 = 0
for item in list3_1:
    y2 = item
    x = 1
    m2 = (y2 - 50) / x
    # print(m2)
    if m2 < 0:  # if a value is less than 0, the yield will be 50
        y2_0 = 50
    else:
        y2_0 = 50 + (m2 * 1)  # Not adjusted fertilizer application
    y2 = 50 + (m2 * 2)  # Adjusted fertilizer application
    print(y2)
    listapp.append(y2)
    sum2_1 = sum2_1 + y2  # Sums all of the results (after process)
    sum2_1_1 = sum2_1_1 + y2_0  # Sums all of the results (before process)

mean = round((sum2_0 + sum2_1) / (len(list3_0) + len(list3_1)))

print("Original sum:" + str(sum))
print("Adjusted sum: " + str(sum2_0_1 + sum2_1_1))
print("Sum after process: " + str(sum2_0 + sum2_1))
print("Mean after process: " + str(mean))


# listfert = matrix.copy(deep=True)
# value = listfert.iterrows()[1]
# for index, row in listfert.iterrows():
#     if value < list3_1[0]:
#         value = 0
#     else:
#         value = 2
# print(listapp)  # Complete list
# for item in list3_0:
#     index = list1.index(item)  # index of item in original, unsorted list

list3 = list2.to_numpy()

k = 0
for item in list3:  # Replaces all values less than 0 with 0
    if item < 0:
        list3[k] = 0
    k += 1

print("Standard Deviation: " + str(np.std(list3)))

kwargs = dict(alpha=0.5, bins=100)
plt.hist(list3, **kwargs, color='g')  # list3 is an array of list2
plt.title("Yield Distribution")
plt.xlabel('Yield')
plt.ylabel('Frequency')
plt.axvline(mean, color='k', linestyle='dashed', linewidth=1)  # this line shows mean on the graph
plt.show()

#  EONR is calculated by the yield increase multiplied by the price of corn. That product is then subtracted by the
#   cost of N

#
