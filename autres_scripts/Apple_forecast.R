library("tseries")
library("zoo")
library("forecast")
library("rugarch")


#get data from yahoo
AAPL.data <-
  get.hist.quote(
    instrument = "AAPL",
    start = "2014-01-01",
    end = "2019-02-25",
    quote = "AdjClose",
    provider = "yahoo",
    compression = "d"
  )


head(AAPL.data)
plot(AAPL.data,
     main = "AAPL closing prices",
     xlab = "Dates",
     ylab = "Prices (USD)")


AAPL.ret = diff(log(AAPL.data)) * 100
AAPL.ret
plot(AAPL.ret,
     main = "Mounthly compound returns",
     xlab = "Dates",
     ylab = "Reterns in percent")


fit1 = auto.arima(AAPL.ret,
                  trace = TRUE,
                  test = "kpss",
                  ic = "bic")
#test arch effect
Box.test(fit1$residuals ^ 2, lag = 12, type = "Ljung-Box")

#specification
res_garch1_spec = ugarchspec(variance.model = list(garchOrder = c(1, 1)),
                             mean.model = list(armaOrder = c(1, 1)))


res_garch1_fit = ugarchfit(spec = res_garch1_spec, data = AAPL.ret)
res_garch1_fit


ctrl = list(tol = 1e-7, delta = 1e-9)
res_garch1_roll = ugarchroll(
  res_garch1_spec,
  AAPL.ret,
  n.start = 120,
  refit.every = 1,
  refit.window = "moving",
  solver = "hybrid",
  calculate.VaR = TRUE,
  VaR.alpha = 0.01,
  keep.coef = TRUE,
  solver.control = ctrl,
  fit.control = list(scale = 1)
)
#report of risks
report(res_garch1_roll,type="VaR",VaR.alpha=0.01, conf.level=0.99)

plot(res_garch1_fit)

#forcasting
res_garch1_fcst=ugarchforecast(res_garch1_fit,n.ahead = 10)


plot(res_garch1_fcst)
res_garch1_fcst

res_garch1_fcst2=ugarchboot(res_garch1_fit,n.ahead = 12, method=c("Partial","Full")[1])
plot(res_garch1_fcst2,which=2)

