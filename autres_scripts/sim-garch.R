#####le premier modele de garch(2,1)
library(fGarch)

spec=garchSpec(model=list(mu=2,omega=.09,alph=c(.15,.3),beta=.4),rseed=9647)
y=garchSim(spec,n=400,extended=TRUE)
y1=y[1:400,1]
plot.ts(y1,xlab="temps")

###############################################

# simuler GARCH(1,1)
# specifier le modèle GARCH(1,1) 

garch11.spec = ugarchspec(variance.model = list(garchOrder=c(1,1)), 
                          mean.model = list(armaOrder=c(0,0)),
                          fixed.pars=list(mu = 0, omega=0.1, alpha1=0.1,
                                          beta1 = 0.7))
set.seed(123)
garch11.sim = ugarchpath(garch11.spec, n.sim=1000)




# utiliser plot pour illustrer les séries simulées et les volatilitées conditionnelles

par(mfrow=c(2,1))
plot(garch11.sim, which=2)
plot(garch11.sim, which=1)
par(mfrow=c(1,1))

par(mfrow=c(3,1))
acf(garch11.sim@path$seriesSim, main="Returns")
acf(garch11.sim@path$seriesSim^2, main="Returns^2")
acf(abs(garch11.sim@path$seriesSim), main="abs(Returns)")
par(mfrow=c(1,1))

par(mfrow=c(2,1))
plot(garch11.sim, which=2)
plot(garch11.sim, which=1)


library(car)
qqPlot(garch11.sim@path$seriesSim, ylab="GARCH(1,1) Returns")
