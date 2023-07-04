Set dict = CreateObject("Scripting.Dictionary")
dict.Add "a","USA"
dict.Add "b","India"
dict.Add "c","UK"
'dict.Add "c","Australia"

print "Total items are -- " & dict.Count ' Size of dictionay

print dict.Exists("a")
print dict.Exists("n")

print dict.Item("b")
print dict.Item("n")
'' All the items in the dictionary
temp = dict.Items
For i=0 to ubound(temp)
print temp(i)
Next

dict.Key("a")="k"
print dict.exists("a")  ' false
print dict.exists("k")   ' true


'' All the keys
temp=dict.Keys
For i=0 to ubound(temp)
print "Key is - " & temp(i) & ". Value is " & dict.Item(temp(i))
Next
















