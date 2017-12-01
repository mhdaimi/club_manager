data = [['500', '192.168.10.100', 'RED'],['500', '192.168.10.200', 'RED'],['64052', '10.10.10.10', 'RED'], ['3802', '192.168.10.10', 'BLUE'], ['488', '10.10.10.10', 'RED'],['488', '10.10.10.10', 'RED'],['500', '192.168.10.10', 'RED']]


key_l=[]
final_list=[]
l_index=[]
my_dict={}
for i in range(len(data)):
    next_val = i+1
    if next_val > len(data):
        break
    k=i+1
    key_l=[]
    for j in data[i+1:]:
        if j[1] == data[i][1]:
            if my_dict:
                if j[1] in my_dict.keys():
                    k=k+1
                    continue
                else:
                    key_l.append(k)
                    k=k+1
                    continue
            else:
                key_l.append(k)
                k=k+1
        else:
            k=k+1
    if key_l:
        key_l.append(data.index(data[i]))
        #make a dictionary with ip as key and values = indexes from list containing this ip 
        my_dict[(data[i][1])] = key_l
         
#iterate over each key and add the 0th value of each list element found at data[index]         
for each_key in my_dict.keys():
    new_val = 0
    for value in my_dict[each_key]:
        new_val += int(data[value][0])                
     
    data[value][0] = new_val
    new_l = data[value]
    final_list.append(new_l)
 
for val in my_dict.values():
    l_index += val
 
#add those list elements which are left
for i in range(len(data)):
    if i in l_index:
        continue
    else:
        final_list.append(data[i])
 
print final_list
 
             
             
