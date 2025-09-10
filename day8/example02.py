# utf8
# author: <NAME>
# date: 2021-01-04
# this is an example for drawing the bar chart
import matplotlib.pyplot as plt
# # get x axis
x = [1,2,3,4,5,6,7,8,9,10]
# get y axis
y= [1,2,3,4,5,6,7,8,9,10]
# plot the bar chart
plt.bar(x,y)
# show the plot
plt.bar(x,y,width=0.5,align='center',color='green')
# show the plot
plt.show()