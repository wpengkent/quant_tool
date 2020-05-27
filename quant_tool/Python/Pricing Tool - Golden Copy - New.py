"""
========================================
Hull White 1 Factor Model Implementation
========================================

==============================================================
Products: Bermudan Swaption and Callable Range Accrual ("CRA")
==============================================================

------------------------------------------
Xlwings Function (to be used in Excel VBA)
------------------------------------------

1) PDE
    Inputs:
        a)direction         - call/put
        b)sign              - Pay/Receive LegA
        c)spotstep          - spotstep
        d)timestep          - timestep
        e)timelabel         - timelabel(items as pde_ts)
        f)mr                - mean reversion
        g)sigma             - hull white calibrated sigma  
        h)payoff_t_a        - Leg A payoff time  
        i)payoff_a          - Leg A payoff
        j)payoff_t_b        - Leg B payoff time  
        k)payoff_b          - Leg B payoff
        l)smartfactor       - Smart multiphication factor (take important date into account)
    
    Returns:
        an array consists of underlying array and option array
     
2) PDE_CalcPeriod
    Inputs:
        a)spotstep          - spotstep
        b)timestep          - timestep
        c)timelabel         - timelabel(items as pde_ts)
        d)mr                - mean reversion
        e)sigma             - hull white calibrated sigma  
        f)periodstart       - CalculationStart (for reset calcperiod each period)
        g)calcperiod        - an array of probability of each timestep
        h)smartfactor       - Smart multiphication factor (take important date into account)
    
    Returns:
        an array consists of cumulative calcperiod (to be used for underlying calculation)
        
3) calcperiod
    Inputs:
        a)spotstep          - spotstep
        b)A                 - hull white formula A (discount)
        c)B                 - hull white formula B (discount)
        d)AF                - hull white formula A (estimation)
        e)BF                - hull white formula B (estimation)
        f)Cal               - yearfrac (discount)
        g)Cal_F             - yearfrac (estimation)
        h)spread            - spread factor for multicurve
        i)upper             - range detail upper
        j)lower             - range detail lower
        k)vol               - effective vol (for payoff dependent smoothing)
        l)rangetype         - Above/Below/Between/Outside
    
    Returns:
        an array consists of probabilities of each time step (non-cumulative)

-----------------
Private Function
-----------------

###Diffusion Process###
    1)pde_matrix - matrix multiphication to solve for diffused value
    2)diag_matrix - to construct diagonal matrix of Pu/Pm/Pd

###Payoff Dependent smoothing###
    1)payoff_dependent_smooth - to smooth the binary function by using lognormal black model

### Payoff Independent Smoothing###
    1)PayoffSmooth_Independent - to identify/smooth the option value based on underlying/option
    2)LinIntp_Payoff_Trapeze - linear interpolation between points
    3)Linear_pts - to find the interval points between two spotstep
    4)smoothval - average of the interval points to find the smooth value

"""

import xlwings as xw
import numpy as np
import time as t
from enum import Enum
from scipy.stats import norm

class pde_ts(Enum):
    call_date=1
    imp_date=2
    crank_nicolson=3
    implict =4

def smoothval(spotstep, int_pts4smoothing, arr_LinIntpUnd, arr_LinIntpOpt, bln_LastUndCF):
    trapeze_const=spotstep*2/10/spotstep/2
    arr_payoffaftersmooth=[None]*int_pts4smoothing
    sumval=0
    
    for i in range(0,int_pts4smoothing):
        if bln_LastUndCF==True:
            arr_payoffaftersmooth[i] = max(arr_LinIntpUnd[i], 0)
            
        else:
            arr_payoffaftersmooth[i] = max(arr_LinIntpUnd[i], arr_LinIntpOpt[i], 0)
        
        arr_trapeze=[None]*int_pts4smoothing
        if i==0 or i==int_pts4smoothing-1:         
            arr_trapeze[i] = 0
            
        else:
            arr_trapeze[i] = arr_payoffaftersmooth[i]
           
        sumval=sumval+arr_payoffaftersmooth[i]+arr_trapeze[i]
  
    output=sumval/2*trapeze_const

    return output
   
def Linear_pts(x,arr_val,int_pts4smoothing):
    upper=arr_val[x+1]
    middle=arr_val[x]
    lower=arr_val[x-1]
    
    int_midpt=(int_pts4smoothing+1)//2
    
    arr_LinIntp=[None]*(int_pts4smoothing)
    arr_LinIntp[0] = (lower + middle) / 2
    arr_LinIntp[int_pts4smoothing-1] = (upper + middle) / 2
    arr_LinIntp[int_midpt-1] = middle
    
    for i in range(1,int_midpt-1):
        arr_LinIntp[i] = arr_LinIntp[i - 1]+ (arr_LinIntp[int_midpt-1] - arr_LinIntp[0]) / (int_midpt - 1)
    
    for i in range(int_midpt,int_pts4smoothing):
        arr_LinIntp[i] = arr_LinIntp[i - 1] + (arr_LinIntp[int_pts4smoothing-1] - arr_LinIntp[int_midpt-1]) / (int_pts4smoothing - int_midpt)
    
    return arr_LinIntp

def LinIntp_Payoff_Trapeze(x,arr_UndVal,int_pts4smoothing,spotstep,arr_OptVal,arr_payoff,int_count,bln_LastUndCF):
    #mid smoothing
    #LinIntp
    arr_LinIntpUnd = Linear_pts(x, arr_UndVal, int_pts4smoothing)
    
    
    if bln_LastUndCF == False: 
        arr_LinIntpOpt = Linear_pts(x, arr_OptVal, int_pts4smoothing)
    else:
        arr_LinIntpOpt=[None]*int_pts4smoothing
        
    #Payoff, Trapeze & Smoothing
    mid_smooth = smoothval(spotstep, int_pts4smoothing, arr_LinIntpUnd, arr_LinIntpOpt, bln_LastUndCF)

    ################################################################################
    
    #adjacent left smoothing
    #LinIntp
    if (x-1)>0:
        arr_LinIntpUnd = Linear_pts(x - 1, arr_UndVal, int_pts4smoothing)
    
    if bln_LastUndCF == False: 
        if (x-1)>0:
            arr_LinIntpOpt = Linear_pts(x - 1, arr_OptVal, int_pts4smoothing)
    else:
        arr_LinIntpOpt=[None]*int_pts4smoothing
        
    #Payoff, Trapeze & Smoothing
    if (x-1)>0:
        left_smooth = smoothval(spotstep, int_pts4smoothing, arr_LinIntpUnd, arr_LinIntpOpt, bln_LastUndCF)
        
    ################################################################################
    
    #adjacent right smoothing
    #LinIntp   
    if (x+1)<(len(arr_UndVal)-1):
        arr_LinIntpUnd = Linear_pts(x + 1, arr_UndVal, int_pts4smoothing)
    
    if bln_LastUndCF == False:
        if (x+1)<(len(arr_OptVal)-1):
            arr_LinIntpOpt = Linear_pts(x + 1, arr_OptVal, int_pts4smoothing)      
    else:
        arr_LinIntpOpt=[None]*int_pts4smoothing
        
    #Payoff, Trapeze & Smoothing
    if (x+1)<(len(arr_OptVal)-1):
        right_smooth = smoothval(spotstep, int_pts4smoothing, arr_LinIntpUnd, arr_LinIntpOpt, bln_LastUndCF)
        
    ################################################################################
    #Replace into Payoff=MAX(Und,Opt,0)

    if max(0,mid_smooth)!=0:
        arr_payoff[x]=mid_smooth

    if (x-1)>0:
        if max(0,left_smooth)!=0:
            arr_payoff[x-1]=max(0,left_smooth)

    if (x+1)<(len(arr_OptVal)-1):
        if max(0,right_smooth)!=0:
            arr_payoff[x+1]=max(0,right_smooth)

    return arr_payoff

def PayoffSmooth_Independent(arr_UndVal,arr_OptVal,spotstep,bln_LastUndCF,bln_Smooth):
    int_pts4smoothing=11
    int_count=len(arr_OptVal)
    
    #Payoff=Max(Und,Opt,0)
    arr_zero=np.zeros_like(arr_OptVal)
    arr_payoff=np.maximum(arr_UndVal,arr_OptVal,arr_zero)
   
    if bln_Smooth == True:
        #Gamma Calc
        arr_down=arr_payoff[:-2]
        arr_base=arr_payoff[1:-1]
        arr_up=arr_payoff[2:]
       
        arr_gamma=arr_down-2*arr_base+arr_up
        arr_gamma/=spotstep**2
        arr_absgamma=np.absolute(arr_gamma)
        
        #Max & Avg Gamma
        max_gamma=np.max(arr_gamma)
        avg_gamma=np.sum(arr_gamma)
        avg_gamma/=int_count-2
       
        #Which spot step to smooth
        condition=0.9*max_gamma+0.1*avg_gamma
        ind_pt2smooth=np.where(arr_absgamma>condition)
    
        #Payoff Smoothing
        for x in ind_pt2smooth[0]:
            arr_payoff=LinIntp_Payoff_Trapeze(x+1,arr_UndVal,int_pts4smoothing,spotstep,arr_OptVal,arr_payoff,int_count,bln_LastUndCF)

    return arr_payoff

def payoff_dependent_smooth(rangetype,upper,lower,forward,vol):
    upper= np.array(upper, dtype=np.float64)
    lower= np.array(lower, dtype=np.float64)
    forward= np.array(forward, dtype=np.float64)
    vol= np.array(vol, dtype=np.float64)
    
    output = np.zeros_like(upper,dtype=np.float64)

    forward = np.maximum(forward,0.0001)
#    (forward>0)*forward + (forward<=0)*0.0001
    
    for i in range(len(rangetype)-1):
        
        if rangetype[i] == "AboveLower":
            output[i,:] = norm.cdf((np.log(forward[i,:]/lower[i,:]*100)-vol[i,:]**2*0.5)/vol[i,:])
    
        elif rangetype[i] == "Outside":
            output[i,:] = np.ones_like(output[i,:])
            output[i,:] -= norm.cdf((np.log(forward[i,:]/lower[i,:]*100)-vol[i,:]**2*0.5)/vol[i,:])
            output[i,:] += norm.cdf((np.log(forward[i,:]/upper[i,:]*100)-vol[i,:]**2*0.5)/vol[i,:])
            
        elif rangetype[i] == "Between":
            output[i,:]  = norm.cdf((np.log(forward[i,:]/lower[i,:]*100)-0.5*vol[i,:]**2)/vol[i,:])
            output[i,:] -= norm.cdf((np.log(forward[i,:]/upper[i,:]*100)-0.5*vol[i,:]**2)/vol[i,:])

        elif rangetype[i] == "BelowUpper":
            output[i,:] = np.ones_like(output[i,:])
            output[i,:] -= norm.cdf((np.log(forward[i,:]/upper[i,:]*100)-vol[i,:]**2*0.5)/vol[i,:])
            
    return output

@xw.func(ret='float64')
def calcperiod(spotstep,A,B,AF,BF,Cal,Cal_F,spread,upper,lower,vol,rangetype):
    
    start = t.time()
    A = np.broadcast_to(A,shape=(len(spotstep),len(A))).transpose()
    B = np.broadcast_to(B,shape=(len(spotstep),len(B))).transpose()
    AF = np.broadcast_to(AF,shape=(len(spotstep),len(AF))).transpose()
    BF = np.broadcast_to(BF,shape=(len(spotstep),len(BF))).transpose()
    Cal = np.broadcast_to(Cal,shape=(len(spotstep),len(Cal))).transpose()
    Cal_F = np.broadcast_to(Cal_F,shape=(len(spotstep),len(Cal_F))).transpose()
    spread = np.broadcast_to(spread,shape=(len(spotstep),len(spread))).transpose()
    upper = np.broadcast_to(upper,shape=(len(spotstep),len(upper))).transpose()
    lower = np.broadcast_to(lower,shape=(len(spotstep),len(lower))).transpose()
    vol = np.broadcast_to(vol,shape=(len(spotstep),len(vol))).transpose()
    
    spotstep = np.broadcast_to(spotstep,shape=(len(A),len(spotstep)))

    HWBond = A*np.exp(-1*B*spotstep)
    HWBond_F = AF*np.exp(-1*BF*spotstep)
    
    forward = (1/HWBond_F-1)/Cal_F+spread
    
    print(str(t.time()-start))
    #output =1 
    start = t.time()
    output = payoff_dependent_smooth(rangetype,upper,lower,forward,vol)*Cal*HWBond
    print(str(t.time()-start))
    return output

def diag_matrix (Pu,Pm,Pd):
    output = (np.diag(Pd,-1)+np.diag(Pm,0)+np.diag(Pu,1))
    return output

'''
To construct a sparse matrix
'''
def sparse_matrix(spotstep,mr,vol,t1,t2,timelabel):
    ds = spotstep[1]-spotstep[0]
    dt = t2-t1
    
    if timelabel ==pde_ts.call_date.value or timelabel==pde_ts.implict.value:
        theta=2
    else:
        theta=1

    Pu= (1/4*dt*(-(vol**2/ds**2) + mr*spotstep/ds))[:-1]
    Pm= dt/2*(vol**2/ds**2+spotstep)
    Pd= (1/4*dt*(-(vol**2/ds**2) - mr*spotstep/ds))[1:]
    
    #boundary condition - Lower Bound
    Pu[0]= 1/2*dt*(mr*spotstep[0]/ds)
    Pm[0]= dt/2*(spotstep[0]-mr*spotstep[0]/ds)
    
    #boundary condition - Upper Bound
    Pm[-1]= dt/2*(spotstep[-1]+mr*spotstep[-1]/ds)
    Pd[-1]= 1/2*dt*(-mr*spotstep[-1]/ds)
    
    i=np.identity(len(spotstep),dtype=np.float64)
    p=theta*diag_matrix (Pu,Pm,Pd)
    a=i+p
    b=i-p
    
    return a,b

def pde_matrix(x,a,b,timelabel):
    if timelabel== pde_ts.call_date.value or timelabel==pde_ts.implict.value:
        #x=np.dot(b,x)
        x=np.linalg.solve(a,x)
    else:
        x=np.dot(b,x)
        x=np.linalg.solve(a,x)
    return x

@xw.func(ret='float64')
def PDE_CalcPeriod(spotstep,timestep,timelabel,mr,sigma,periodstart,calcperiod,smartfactor):
    #Transform to array instead of list
    periodstart = np.array(periodstart,dtype=np.float64)
    calcperiod = np.array(calcperiod,dtype=np.float64)
    sigma = np.array(sigma,dtype=np.float64)
    smartfactor = np.array(smartfactor,dtype=np.float64)
    timestep = np.array(timestep,dtype=np.float64)
    timelabel = np.array(timelabel,dtype=np.float64)
    spotstep=np.array(spotstep,dtype=np.float64)
    mr=np.float64(mr)
    
    pde_temp = np.zeros(len(spotstep),dtype=np.float64)
    output = np.zeros((len(timestep),len(spotstep)),dtype=np.float64)
    cnt=0
    start = t.time()
    #sum_x = 0
    
    vol_check  = 0
    dt_check = 0
    timelabel_check=0
    
    for i in range(len(timestep)-2,-1,-1):
        t1 = timestep[i]
        t2 = timestep[i+1]
        dt = t2-t1
        vol =sigma[np.searchsorted(sigma[:,1],t2),0]
        pde_temp = output[i+1,:]
        
        if vol!=vol_check or dt!=dt_check or timelabel[i+1]!=timelabel_check:
            a,b = sparse_matrix(spotstep,mr,vol,t1,t2,timelabel[i+1])
            vol_check =vol
            dt_check = dt
            timelabel_check =timelabel[i+1]
            
        if timestep[i+1]==periodstart[cnt]:
            output[i,:]=calcperiod[i,:]
            if cnt < (len(periodstart)-1):
                cnt=cnt+1
        else:
            pde_temp = pde_matrix(pde_temp,a,b,timelabel[i+1])
            output[i,:]=(pde_temp*smartfactor[i]+calcperiod[i,:])
            
    print("CalcPeriod Total:"+ str((t.time()-start)))
    #print("CalcPeriod Diffusion:" + str(sum_x))
    return output

@xw.func(ret='float64')
def PDE(direction,sign,spotstep,timestep,timelabel,mr,sigma,payoff_t_a,payoff_a,payoff_t_b,payoff_b,smartfactor,bln_Smooth):
    
    bln_LastUndCF = True
    #Transform to array instead of list
    payoff_a = np.array(payoff_a,dtype=np.float64)
    payoff_b = np.array(payoff_b,dtype=np.float64)
    payoff_t_a = np.array(payoff_t_a,dtype=np.float64)
    payoff_t_b = np.array(payoff_t_b,dtype=np.float64)
    sigma = np.array(sigma,dtype=np.float64)
    smartfactor = np.array(smartfactor,dtype=np.float64)
    timestep = np.array(timestep,dtype=np.float64)
    timelabel = np.array(timelabel,dtype=np.float64)
    spotstep=np.array(spotstep,dtype=np.float64)
    mr=np.float64(mr)

    undval_output = np.zeros((len(timestep),len(spotstep)),dtype=np.float64)
    opt_output = np.zeros((len(timestep),len(spotstep)),dtype=np.float64)
    
    ds = spotstep[-1]-spotstep[-2]
    
    '''
    Diffusion Process Starts
    '''
    cnt_a, cnt_b=0,0
    undval_temp = np.zeros(len(spotstep))
    opt_temp = np.zeros(len(spotstep))
    
    ''''
    Underlying first step
    '''
    if timestep[-1] in payoff_t_a:
        undval_temp += direction*-sign*payoff_a[:,cnt_a]
        cnt_a+=1
    if timestep[-1] in payoff_t_b:
        undval_temp += direction*sign*payoff_b[:,cnt_b]
        cnt_b+=1
        
    undval_output[-1,:]=undval_temp*smartfactor[-1]
    
    '''
    Option first step
    '''
    if timelabel[-1]==pde_ts.call_date.value:
        #opt_temp = np.maximum(opt_temp,undval_output[-1,:])
        opt_output[-1,:]=PayoffSmooth_Independent(undval_output[-1,:],opt_temp*smartfactor[-1],ds,bln_LastUndCF,bln_Smooth)
        #[-1,:]=opt_temp*smartfactor[-1]

    else:
        opt_temp = np.zeros(len(spotstep))
        opt_output[-1,:]=opt_temp
     
    start = t.time()
    
    '''
    Underlying and Option Subsequent Step
    '''
    vol_check  = 0
    dt_check = 0
    timelabel_check=0
    
    for i in range(len(timestep)-2,-1,-1):
        t1 = timestep[i]
        t2 = timestep[i+1]
        dt = t2-t1
        vol =sigma[np.searchsorted(sigma[:,1],t2),0]
        
        if vol != vol_check or dt != dt_check or timelabel[i+1] != timelabel_check:
            a,b = sparse_matrix(spotstep,mr,vol,t1,t2,timelabel[i+1])
            vol_check =vol
            dt_check = dt
            timelabel_check =timelabel[i+1]
            
        undval_temp = pde_matrix(undval_output[i+1,:],a,b,timelabel[i+1])
        opt_temp = pde_matrix(opt_output[i+1,:],a,b,timelabel[i+1]) 
        
        if timestep[i] in payoff_t_a:
            undval_temp += direction*-sign*payoff_a[:,cnt_a]
            if cnt_a < (len(payoff_a[0,:])-1):
                cnt_a+=1
                
        if timestep[i] in payoff_t_b:
            undval_temp += direction*sign*payoff_b[:,cnt_b]
            if cnt_b < (len(payoff_b[0,:])-1):
                cnt_b+=1

        undval_output[i,:]=undval_temp*smartfactor[i]
                
        if timelabel[i]==pde_ts.call_date.value:
            #opt_output[i,:]=np.maximum(undval_output[i,:],opt_temp*smartfactor[i])
            opt_output[i,:]=PayoffSmooth_Independent(undval_output[i,:],opt_temp*smartfactor[i],ds,bln_LastUndCF,bln_Smooth)
            bln_LastUndCF=False
        else:
            opt_output[i,:]=opt_temp*smartfactor[i]
            
    print("Und Total:"+ str((t.time()-start)))
    #print("Und Diffusion:" + str(sum_x))
    
    #opt_output = PDE_Opt(spotstep,timestep,timelabel,mr,sigma,undval_output,smartfactor)
    
    final_output = np.stack((-direction*undval_output,opt_output))
#    wb = xw.books.active
#    wb.app.activate(steal_focus=True)
    return final_output

