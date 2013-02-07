import scipy as sp
import numpy as np
import matplotlib as mpl


def threeP(y, x,  div):
    
    # Calculate minimum and maximum values for X axis
    xMin = np.min(x)
    xMax = np.max(x)
    
    # Calculate  interval for 
    dx = (xMax - xMin) / div
    
    mult = np.linspace(1,  div,  div)
    
    r2 = np.zeros(len(x))
    
    # First Pass
    
    for m in mult:
        xTemp = x - (xMin + (m-1)*dx)
        xTemp[xTemp < 0.] = 0.
        A = np.vstack([xTemp,  np.ones(len(xTemp))]).T
        
        model, residual = np.linalg.lstsq(A, y)[0]
        
        r2[m] = 1 - residual/(y.size * y.var())
        
        
    print "Model: ",  model
    print "Residual: ",  residual
    print "r2: ",  r2
    
    # Second Pass
    
    
    
def main():
    print "Main"

    E = np.array([3638.71,  3755.17,  3416.13,  3730.00,  4127.59,  4453.13,  5160.00,  5617.24,  5420.69,  4706.25,  4096.55,  3635.29,  3640.00,  3806.90, 3764.52, 3724.14, 4340.00, 4837.50, 5340.00, 5565.52, 5271.43, 4781.82, 4168.97])
    T = np.array([-99., 35.78, 36.55, 56.55, 62.26, 70.14, 80.76, 82.53, 79.95, 71.75, 60.00, 47.83, 34.26, 32.54, 42.14, 47.29, 59.62, 70.51, 79.10, 84.60, 78.51, 68.71, 58.43])

    threeP(E,  T,  10)
     
if __name__=="__main__":
    main()
