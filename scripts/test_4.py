from future.utils import isint
data = ["02-sep-2016, blah blah, blah, 83838338",2, 4,6,1,2,"02-sep-2016, blah blah, blah, 83838338",3,0,0,"03-Aug-2000, blah, 300033","03-Aug-2000, blah, 300033"]

vals=[]
final_data = "%d,%s"
formatted_rec = []
for each_val in data:
    
    if not isint(each_val) and "-" in each_val:
        if vals:
            max_digit = max(vals)
        else:
            max_digit = 0
        vals=[]
        formatted_rec.append(final_data %(max_digit, each_val))
    else:
        vals.append(each_val)
        
for each_rec in formatted_rec:
    print each_rec    