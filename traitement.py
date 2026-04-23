def puissance(a,n):
    pow=1
    i=1
    while(i!=n+1):
        pow=pow*a
        i=i+1
    print(pow)
    return pow

puissance(2,7)

def factoriel(n):
    if(n==0 or n==1):
        return(n)
    else:
        return n*factoriel(n-1)

r=factoriel(5)
print(r)  
    
        