x = [66,1,66,3,66,55]
while x:
    list1 = [x.pop(x.index(max(x)))]
print(list1)
print(x)