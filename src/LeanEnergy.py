import scipy as sp
import numpy as np
import matplotlib.pyplot as plt


def threeP(y, x,  div):
    print "...First Pass..."
    
    # Calculate minimum and maximum values for X axis
    xMin = np.min(x)
    xMax = np.max(x)    

    print "xMin: ", xMin
    print "xMax: ", xMax
    print "y mean: ", np.mean(y)
    
    # Calculate  interval for 
    dx = (xMax - xMin) / div
    print "dx: ", dx
    
    #np.linespace(start, stop, number of samples, include endpoint or not)

    #mult = np.linspace(1,  div,  num=div,  endpoint=True)
    mult = np.linspace(1,  div,  num=(div-1),  endpoint=False)
    
    r2 = np.zeros(len(x))
    adjr2 = np.zeros(len(x))
    rmse = np.zeros(len(x))
    cvrmse = np.zeros(len(x))
    yHat = np.zeros(len(x))

    # First Pass
    
    for m in mult:

        xTemp = x

        # Calculate various change point temperatures to test.  Not sure if I should include m or (m-1)?  Will need to research this further.  If including m, then need to set "endpoint = True" in "np.linspace" above.
        xCp = xMin + m*dx
        # xCp = xMin + (m-1)*dx

        #  THis is where we would test for 3PC or 3PH models using a control statement.  i.e. If 3PH then do ...  elseif 3PC then do...
        #  Depending on the choice, you would alter the way that xCp for 0 terms.  For now, we are only considering 3PC.

        xTemp = xTemp - xCp
        xTemp[xTemp <= 0.] = 0.
        
        print "Round: ",  m
        print "xCp: ",  xCp
        print "xTemp: ",  xTemp
        
        # In the future, this is where I would add additional x variables for multivariate analysis.  May need a control structure
        # to manage numpy and scipy implementations of OLS regression.  As of current, it is strictly an univariate OLS regression.
        
        A = np.vstack([xTemp,  np.ones(len(xTemp))]).T

        result1 = np.linalg.lstsq(A, y)

        # Need to calculate RMSE, CV-RMSE, r2, adjR2, sigma, roe
        # yHat[np.where(xTemp == 0.)] = np.mean(y[np.where(xTemp == 0.)])
        # yHat[np.where(xTemp != 0.)] = result1[0][0]*xTemp[np.where(xTemp != 0.)] + result1[0][1]
        # ssm:  np.sum(np.square(y - np.mean(y)))
        # sse: np.sum(np.square(y - yHat))
        # mse (kelly): sse / (number of rows - number of columns)
        # mse (wikipedia): np.sum(np.square(yHat - y)) / number of rows
        # r2:  1 - sse / ssm
        # rmse: np.sqrt(mse)
        # cv-rmse: rmse / ymean * 100

        yHat[np.where(xTemp == 0.)] = np.mean(y[np.where(xTemp == 0.)])
        yHat[np.where(xTemp != 0.)] = result1[0][0]*xTemp[np.where(xTemp != 0.)] + result1[0][1]
         
        ssm = np.sum(np.square(y - np.mean(y)))
        sse = np.sum(np.square(y-yHat))
        mse = np.sum(np.square(yHat - y)) / len(y)
        rmse[m] = np.sqrt(mse)
        cvrmse[m] = rmse[m] / np.mean(y)
        r2[m] = 1 - sse/ssm
        adjr2[m] = 1 - (len(y) - 1) / (len(y) - (A.shape[1] -1))
     
        print "yHat: ", yHat
        print "rmse: ",  rmse[m]
        print "cv-rmse: ",  cvrmse[m]
        print "r2: ",  r2[m]
        print "adj-R2: ",  adjr2[m]

        # Now need to select where rmse is lowest and use that as the starting point for the next iteration of the grid search.

        plot(x,  y,  yHat)
        
        """
        #print "r2: ", r2 = 1 - sse / ssm
        
        
        
        r2[m] = rsquared(result1[1],  y)

    print "m: ", result1[0][0]
    print "b: ", result1[0][1]
    
    #print "r2: ",  r2

    mIndex = np.mean(np.median(np.where(r2 == np.max(r2))))

    
    
    # Second Pass

    print "...Second Pass..."
    
    xMax = x[mIndex] + dx
    xMin = x[mIndex] - dx
    
    print "xMax: ",  xMax
    print "xMin: ",  xMin
    
    dx = (xMax - xMin) / div
    
    mult = np.linspace(1,  div,  div)
    
    r2 = np.zeros(len(x))


    
    for m in mult:
        xTemp = x - (xMin + (m-1)*dx)
        xTemp[xTemp < 0.] = 0.
        A = np.vstack([xTemp,  np.ones(len(xTemp))]).T
        
        result2 = np.linalg.lstsq(A, y)
        
        #r2[m] = rsquared(result2[1],  y)

    print "m: ", result2[0][0]
    print "b: ", result2[0][1]
    
    #print "r2: ",  r2
    #print "b3: ", xMin + np.mean(np.median(np.where(r2 == np.max(r2))))*dx
    
    """
    
def rsquared(r,  y):
    return 1 - r/(y.size * y.var())

def plot(x,  y,  yhat):
    plt.plot(x, y, 'o')
    plt.plot(x, yhat, 'r')
    plt.show()

def slidingThreeP(x, y):
    # Number of runs for monthly data; need to write an algorithm to detect whether hourly, daily, monthly.
    runs = x.length - 12
    
def main():
    print "Main"

    #E = np.array([3638.71,  3755.17,  3416.13,  3730.00,  4127.59,  4453.13,  5160.00,  5617.24,  5420.69,  4706.25,  4096.55,  3635.29,  3640.00,  3806.90, 3764.52, 3724.14, 4340.00, 4837.50, 5340.00, 5565.52, 5271.43, 4781.82, 4168.97])
    #T = np.array([-99., 35.78, 36.55, 56.55, 62.26, 70.14, 80.76, 82.53, 79.95, 71.75, 60.00, 47.83, 34.26, 32.54, 42.14, 47.29, 59.62, 70.51, 79.10, 84.60, 78.51, 68.71, 58.43])

    E = np.array([3755.17,  3416.13,  3730.00,  4127.59,  4453.13,  5160.00,  5617.24,  5420.69,  4706.25,  4096.55,  3635.29,  3640.00,  3806.90, 3764.52, 3724.14, 4340.00, 4837.50, 5340.00, 5565.52, 5271.43, 4781.82, 4168.97])
    T = np.array([35.78, 36.55, 56.55, 62.26, 70.14, 80.76, 82.53, 79.95, 71.75, 60.00, 47.83, 34.26, 32.54, 42.14, 47.29, 59.62, 70.51, 79.10, 84.60, 78.51, 68.71, 58.43])

    threeP(E,  T,  10)
     
if __name__=="__main__":
    main()
