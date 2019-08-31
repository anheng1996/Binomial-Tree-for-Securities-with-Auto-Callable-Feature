import math
import numpy as np
import xlwt
from matplotlib import pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import time

start = time.clock()

#####parameters
S=2704.1
threshold=1757.665
r=0.025263132
T=3


####CRR
def CRR(sigma,div,N):
    dt=T/N
    u=math.exp(sigma*(math.sqrt(dt)))
    d=1/u
    p1=(math.exp((r-div)*dt)-d)/(u-d)
    p2=1-p1
    
    ####Stock Price Matrix
    stock=np.ones((N+1,N+1))*0
    stock[0,0]=S
    for i in range(1,N+1):
        stock[i,0]=stock[i-1,0]*d
    for i in range(1,N+1):
        for j in range(i,N+1):
            stock[j,i]=u*stock[j-1,i-1]
            
    ####Option Matrix
    option=np.ones((N+1,N+1))*0
    for i in range(0,N+1):
        if stock[N,i]>=S:
            option[N,i]=1300
        elif threshold<=stock[N,i]<S:
            option[N,i]=1000
        else:
            option[N,i]=1000*stock[N,i]/S
    for i in range(N-1,-1,-1):
        for j in range(0,i+1):
            if i==math.floor(2/dt):
                if stock[i,j]>=S:
                    option[i,j]=1200*math.exp(r*(i*dt-2-3/365))
                else:
                    option[i,j]=math.exp(-r*dt)*(p1*option[i+1,j+1]+p2*option[i+1,j])
            elif i==math.floor(1/dt):
                if stock[i,j]>=S:
                    option[i,j]=1100*math.exp(r*(i*dt-1-3/365))
                else:
                    option[i,j]=math.exp(-r*dt)*(p1*option[i+1,j+1]+p2*option[i+1,j])
            else:
                option[i,j]=math.exp(-r*dt)*(p1*option[i+1,j+1]+p2*option[i+1,j])
    
    return(option[0,0])

work_book=xlwt.Workbook(encoding='utf-8')
sheet=work_book.add_sheet('project1')
sheet.write(0,0,"N")
sheet.write(0,1,"value")

####output
for i in range(51,10002,50):
    sheet.write(int((i-1)/50),0,i)
    sheet.write(int((i-1)/50),1,CRR(sigma=0.18921,div=0.0212,N=i))
    work_book.save('project1.xls')



#####sensitivity analysis
fig = plt.figure()
ax = Axes3D(fig)
X = np.arange(0.16, 0.27, 0.01)    ####sigma span
Y = np.arange(0.018, 0.033, 0.001)    ####div yield span
X, Y = np.meshgrid(X, Y)
Z=np.ones((16,12))*0
for i in range(0,12,1):
    for j in range(0,16,1):
        Z[j-1,i-1]=CRR(sigma=0.16+0.01*i,div=0.018+0.001*j,N=10001)
ax.plot_surface(X, Y, Z,cmap="rainbow",alpha=0.5)
ax.set_xlabel('sigma', color='b')
ax.set_ylabel('div', color='b')
ax.set_zlabel('price', color='b')
ax.set_zlim(900, 1050)
#ax.text(0.1,0.02,CRR(sigma=0.18921,div=0.0212,N=10001),"our estimation")
plt.show()


elapsed = (time.clock() - start)
print("Time used:",elapsed)