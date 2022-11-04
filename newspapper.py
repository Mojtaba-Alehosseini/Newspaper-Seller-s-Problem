import random


def day_status():
    P = random.random()
    #print(P)
    if(0 <= P <= 0.35):
        return('Good')
    elif(0.35 < P <= 0.80):
        return('Fair')
    else:
        return('Poor')

def sale(day_):
    P = random.random()
    #print(P)
    
    if day_ == 'Good':
        if(P <= 0.03):
            return(40)
        elif(P <= 0.08):
            return(50)
        elif(P <= 0.23):
            return(60)
        elif(P <= 0.43):
            return(70)
        elif(P <= 0.78):
            return(80)
        elif(P <= 0.93):
            return(90)
        else:
            return(100)
    elif day_ == 'Fair':
        if(P <= 0.10):
            return(40)
        elif(P <= 0.28):
            return(50)
        elif(P <= 0.68):
            return(60)
        elif(P <= 0.88):
            return(70)
        elif(P <= 0.96):
            return(80)
        else:
            return(90)
        
    elif day_ == 'Poor':
        if(P <= 0.44):
            return(40)
        elif(P <= 0.66):
            return(50)
        elif(P <= 0.82):
            return(60)
        elif(P <= 0.94):
            return(70)
        else:
            return(80)

def Q_Newspapers_sale(Q,days):
    Sum = 0
    for day in range(0,days):

        TOND = day_status()
        Demand = sale(TOND)

        if Demand <= Q:    
            RFS = Demand * 0.50
            LPFED = 0
            SFSOS = (Q - Demand) * 0.05
        else:
            RFS = Q * 0.50
            LPFED = (Demand - Q) * 0.17
            SFSOS = 0

        Daily_cost = Q * 0.33
        Daily_profit = RFS - Daily_cost + LPFED + SFSOS
        Sum += Daily_profit
        #print(TOND,Demand,round(RFS,2),round(LPFED,2),round(SFSOS,2),Daily_cost,round(Daily_profit, 2))
    return(round(Sum, 2))

for Q in [50,60,70,80,90,100]:
    Sum_ = 0
    for _ in range(0,500):
        profit = Q_Newspapers_sale(Q,60)
        Sum_ += profit
    print('Profit in 60 Day with ',Q,' is :',round(Sum_/500, 2))