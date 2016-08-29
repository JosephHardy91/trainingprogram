__author__ = 'Administrator'

def fibloop(x):
    inputs=[int(y) for y in x.split('\n')]
    inputs=inputs[1:]
    answers=[]
    l=0
    z=0
    for input in inputs:
        print input
        while z <= input and input!=0:
            #print x-z
            if z==0:
                j=0
                k=1
            else:
                k=j
                j=z
            z=j+k
            l+=1
        answers.append(l-1)
    for answer in answers:
        print answer,
x=raw_input()
inputs=[int(y) for y in x.split('\n')]
print inputs
fibloop(raw_input())