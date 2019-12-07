library(xlsx)
library(base)
library(car)
#statistiques descriptives
library(pastecs)
library(rugarch)
library(tseries)
library(aTSA)
library(astsa)
library(grDevices)
library(forecast)
fn="./Data.xlsx"
#--------------------------- nettoyage de donnÃ©es -----------------------------------
cleaning_data<-function(filename,sheetindex){
  #reading excel sheet into a data.frame
  clean_data=read.xlsx2(filename,sheetIndex=sheetindex);
  #converting from factor to numeric and removing missing data
  clean_data[,1:ncol(clean_data)]=sapply(clean_data[,1:ncol(clean_data)],as.character);
  clean_data[,2:ncol(clean_data)]=sapply(clean_data[,2:ncol(clean_data)],as.numeric);
  clean_data=clean_data[complete.cases(clean_data), ];
}
#--------------------------- calcule des statistiques descriptives -----------------------------------
descStat<-function(filename=fn){
  for(i in 1:5){
    data<-cleaning_data(fn,i)
    desc_stat<-t(stat.desc(data[,2:ncol(data)])[c(8,9,12,13),])
    write.xlsx(file="Statistique_descriptive.xlsx",desc_stat,sheetName=paste("variable_",i,sep=""),append=TRUE) 
  }
  print("Pour les statistiques descriptives veuillez voir le fichier Excel Statistique_descriptive.xlsx")
}
#--------------------------- teste de normalitÃ© -----------------------------------
normality_test<-function(filename=fn){
  for(i in 1:5){
    data<-0
    data<-cleaning_data(fn,i)
    test=0
    w=0
    p=0
    periodes=""
    for (j in 2:ncol(data)){
      w[j-1]<-shapiro.test(data[[j]])[[1]]
      p[j-1]<-shapiro.test(data[[j]])[2]
      if(p[j-1]<=0.05){
        test[j-1]<-"non gaussien"
      }else{
        test[j-1]<-"gaussien"
      }
    }
    periodes=colnames(data[,2:ncol(data)])
    result<-cbind(periodes,w,p,test)
    colnames(result)<-c("Periode","Statistique de test","P-valeur","test normalite")
    write.xlsx2(file="test_normalite.xlsx",result,sheetName=paste("variable_",i,sep=""),append=TRUE,row.names = FALSE) 
  }
  print("Pour les rÃ©sultats de test de normalitÃ© veuillez voir le fichier Excel test_normalite.xlsx")
}
#--------------------------- teste de correlation -----------------------------------
correlation_test<-function(filename=fn){
  for(i in 1:5){
    data<-cleaning_data(fn,i)
    corr_matrix<-cor(data[,2:ncol(data)])
    write.xlsx(file="matrice_correlation.xlsx",corr_matrix,sheetName=paste("variable_",i,sep=""),append=TRUE) 
  }
  print("Pour les resultats de test de correlation veuillez voir le fichier Excel matrice_correlation.xlsx")
}

#--------------------------- test d'existence de l'effet arch pour une periode -----------------------------------
arch_effect_exist<-function(periode){
  p<-Box.test(periode^2,lag=1,type="Ljung-Box")[[3]]
  if(p>0.05){
    return(FALSE)
  }else{
    return(TRUE)
  }
}

#--------------------------- test de stationaritÃ© d'une pÃ©riode -----------------------------------
is_stationary<-function(periode){
  p<-as.numeric(adf.test(periode,nlag=1,output = FALSE)$type1[1,3])
  if(p>0.05){
    return(FALSE)
  }else{
    return(TRUE)
  }
}
#--------------------------- test d'existence de l'effet ARCH pour tout les donnÃ©es -----------------------------------


arch_effect_test<-function(filename=fn){
  for(i in 1:5){
    data<-0
    data<-cleaning_data(fn,i)
    test1=""
    test2=""
    test3=""
    x_sq=0
    p_stat=0
    p_ar_eff=0
    periodes=""
    model="";
    for (j in 2:ncol(data)){
      p_ar_eff[j-1]<-Box.test(data[,j]^2,lag=1,type="Ljung-Box")[[3]]
      p_stat[j-1]<-as.numeric(adf.test(data[,j],nlag=1,output = FALSE)$type1[1,3])
      if(p_ar_eff[j-1]>0.05){
        test1[j-1]<-"non auto-correlee"
        model[j-1]<-"ARIMA/SARIMA"
      }else{
        test1[j-1]<-"auto-correlee"
        model[j-1]<-"ARCH/GARCH"
      }
      if(p_stat[j-1]>0.05){
        test2[j-1]<-"non stationnaire" 
      }else{
        test2[j-1]<-"stationnaire" 
      }
    }
    periodes=colnames(data[,2:ncol(data)])
    result<-cbind(periodes,p_ar_eff,test1,p_stat,test2,model)
    colnames(result)<-c("Periode","P-valeur","Test auto-correlation","P-valeur","Test stationarite","Modele")
    write.xlsx2(file="test_effet_arch_stationarite.xlsx",result,sheetName=paste("variable_",i,sep=""),append=TRUE,row.names = FALSE) 
  }
  print("Pour les rÃ©sultats de test d'existence d'effet arch  veuillez voir le fichier Excel test_effet_arch_stationarite.xlsx")
}


#--------------------------- models ARCH/GARCH -----------------------------------
arch_garch<-function(periode,periode_name,nb_pred){
  a1a0.spec = ugarchspec(variance.model = list(garchOrder=c(1,0)),mean.model = list(armaOrder=c(0,0)));
  g11a0.spec = ugarchspec(variance.model = list(garchOrder=c(1,1)),mean.model = list(armaOrder=c(0,0)));
  a1a1.spec = ugarchspec(variance.model = list(garchOrder=c(1,0)),mean.model = list(armaOrder=c(1,0)));
  g11a1.spec = ugarchspec(variance.model = list(garchOrder=c(1,1)),mean.model = list(armaOrder=c(1,0)));
  
  a1a0.fit = ugarchfit(spec=a1a0.spec, data=periode,solver.control=list(trace = 1));
  g11a0.fit = ugarchfit(spec=g11a0.spec, data=periode,solver.control=list(trace = 1));
  a1a1.fit = ugarchfit(spec=a1a1.spec, data=periode,solver.control=list(trace = 1));
  g11a1.fit = ugarchfit(spec=g11a1.spec, data=periode,solver.control=list(trace = 1));
  ica1a0=infocriteria(a1a0.fit)[1];
  icg11a0=infocriteria(g11a0.fit)[1];
  ica1a1=infocriteria(a1a1.fit)[1];
  icg11a1=infocriteria(g11a1.fit)[1];
  
  A=c(ica1a0,icg11a0,ica1a1,icg11a1);
  
  cmp=A[1];
  k=1;
  tmp=1;
  while(k<4){
    if(cmp>A[k]){
      cmp=A[k];
      tmp=k;
    }
    k=k+1;
  }
  forecast_value<-ugarchforecast(switch(tmp,a1a0.fit,g11a0.fit,a1a1.fit,g11a1.fit),n.ahead=nb_pred)
  plot(forecast_value,which=1,title=NULL)
  title(paste("resultats de prdiction de la ",periode_name," avec GARCH"))
  return(forecast_value)
  #colnames(predicted_serie)<-colnames(data[,2:ncol(data)])
  #write.xlsx2(file="prediction.xlsx",predicted_serie,sheetName=paste("variable_",i,sep=""),append=TRUE,row.names = FALSE)
}
#--------------------------- models ARCH/GARCH -----------------------------------

sarimaFit <- function(periode,periode_name,nb_pred=nb_pred){
  #matrice de correlation : pour etudier le degr? dindependence
  
  sf1=SF(periode,n.ahead=nb_pred,1,1,1,0,1,1,12,xname=periode_name);
  sf2=SF(periode,n.ahead=nb_pred,0,1,1,0,1,1,12,xname=periode_name);
  sf3=SF(periode,n.ahead=nb_pred,1,1,1,0,1,0,12,xname=periode_name);
  sf4=SF(periode,n.ahead=nb_pred,0,1,1,0,1,1,12,xname=periode_name);
  
  b<-matrix(,4,7);
  b[1,]<-c(1,1,1,0,1,1,12)
  b[2,]<-c(0,1,1,0,1,1,12)
  b[3,]<-c(1,1,1,0,1,0,12)
  b[4,]<-c(0,1,1,0,1,1,12)
  A=c(sf1[[2]],sf2[[2]],sf3[[2]],sf4[[2]]);
  ind=which.min(A);
  co=switch(as.numeric(ind),sf1,sf2,sf3,sf4);
  #fc=as.vector(co[[1]]$pred);
  sf1=SF(periode,n.ahead=nb_pred,b[ind,1],b[ind,2],b[ind,3],b[ind,4],b[ind,5],b[ind,6],b[ind,7],xname="Data",output = TRUE,title=paste("resultats de prdiction de la ",periode_name," en appliquant SARIMA"));
  print("*******************************")
  print(str(co))
  return(co)
}

SF <- function(xdata,n.ahead,p,d,q,P=0,D=0,Q=0,S=-1,tol=sqrt(.Machine$double.eps),no.constant=FALSE, plot.all=FALSE,
               xreg = NULL, newxreg = NULL,xname,output=FALSE,title="graph de prediction"){ 
  #
  layout = graphics::layout
  par = graphics::par
  plot = graphics::plot
  grid = graphics::grid
  box = graphics::box
  abline = graphics::abline
  lines = graphics::lines
  frequency = stats::frequency
  na.pass = stats::na.pass
  as.ts = stats::as.ts
  ts.plot = stats::ts.plot
  xy.coords = grDevices::xy.coords
  #
  #xname=deparse(substitute(xdata))
  xdata=as.ts(xdata) 
  n=length(xdata)
  if (is.null(xreg)) {  
    constant=1:n
    xmean = rep(1,n);  if(no.constant==TRUE) xmean=NULL
    if (d==0 & D==0) {
      fitit=stats::arima(xdata, order=c(p,d,q), seasonal=list(order=c(P,D,Q), period=S),
                         xreg=xmean,include.mean=FALSE, optim.control=list(reltol=tol));
      nureg=matrix(1,n.ahead,1);  if(no.constant==TRUE) nureg=NULL          
    } else if (xor(d==1, D==1) & no.constant==FALSE) {
      fitit=stats::arima(xdata, order=c(p,d,q), seasonal=list(order=c(P,D,Q), period=S),
                         xreg=constant,optim.control=list(reltol=tol));
      nureg=(n+1):(n+n.ahead)       
    } else { fitit=stats::arima(xdata, order=c(p,d,q), seasonal=list(order=c(P,D,Q), period=S), 
                                optim.control=list(reltol=tol));
    nureg=NULL   
    }
  }   
  if (!is.null(xreg)) {
    fitit = stats::arima(xdata, order = c(p, d, q), seasonal = list(order = c(P, 
                                                                              D, Q), period = S), xreg = xreg)
    nureg = newxreg		
  }
  
  #--
  fore <- stats::predict(fitit, n.ahead, newxreg=nureg)  
  #-- graph:
  U  = fore$pred + 2*fore$se
  L  = fore$pred - 2*fore$se
  U1 = fore$pred + fore$se
  L1 = fore$pred - fore$se
  if(plot.all)  {a=1} else  {a=max(1,n-100)}
  minx=min(xdata[a:n],L)
  maxx=max(xdata[a:n],U)
  t1=xy.coords(xdata, y = NULL)$x 
  if(length(t1)<101) strt=t1[1] else strt=t1[length(t1)-100]
  t2=xy.coords(fore$pred, y = NULL)$x 
  endd=t2[length(t2)]
  if(plot.all)  {strt=time(xdata)[1]} 
  xllim=c(strt,endd)
  par(mar=c(2.5, 2.5, 1, 1), mgp=c(1.6,.6,0))
  if(output){
    ts.plot(xdata,fore$pred, type="n", xlim=xllim, ylim=c(minx,maxx), ylab=xname)
    title(title);
    grid(lty=1, col=gray(.9)); box()
    if(plot.all) {lines(xdata)} else {lines(xdata, type='o')}   
    #  
    xx = c(time(U), rev(time(U)))
    yy = c(L, rev(U))
    polygon(xx, yy, border=8, col=gray(.6, alpha=.2) ) 
    yy1 = c(L1, rev(U1))
    polygon(xx, yy1, border=8, col=gray(.6, alpha=.2) ) 
    
    lines(fore$pred, col="red", type="o")
    #legend(x=0,y=NULL,legend=c("valeurs predites","valeurs originale"),col = c("red","gray"))
  }
  fr=list(fore,fitit$aic,sqrt(fitit$sigma2));
  #> a$value[[2]]
  return(fr)
}
#--------------------------- fonction de predition -----------------------------------
perdiction<-function(filename=fn){
  dev.new();
  dev.hold();
  for(i in 1:5){
    data<-0
    data<-cleaning_data(fn,i)
    p=0
    periodes=""
    last_year<-as.numeric(substr(data[nrow(data),1],start=1,stop=4))+2
    first_year<-as.numeric(substr(data[1,1],start=1,stop=4))
    nb_pred<-length(last_year:2031)
    nrws<-nrow(data)+nb_pred;
    predicted_series<-matrix(,nrws,ncol(data[,2:ncol(data)]));
    pdf(file=paste("Varriable_",i,"plots.pdf",sep=""),title="")
    if(nrow(data)>=14){
      for (j in 2:ncol(data)){
        periode<-ts(data[,j],start=first_year,frequency=1)
        diff_periode<-round(diff(periode),3)
        if(arch_effect_exist(periode)){
          #l'effet arch existe donc on utilse le model arch/garch pour la prediction
          if(is_stationary(periode)){
            forecasted_value<-arch_garch(periode,colnames(data)[j],nb_pred=nb_pred)
            ###################we use 95% prediction interval
            fserie=ts(forecasted_value@forecast$seriesFor[,1])
            sigma_residuals=sd(forecasted_value@model$modeldata$residuals)
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            #################"""
            predicted_series[,j-1]<-c(as.vector(periode),pre)
            }else{
            forecasted_value<-arch_garch(diff_periode,paste("diff(",colnames(data)[j],")"),nb_pred=nb_pred)
            ###################"
            fserie=ts(forecasted_value@forecast$seriesFor[,1])
            sigma_residuals=sd(forecasted_value@model$modeldata$residuals)
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            #################"""
            predicted_series[2:nrow(predicted_series),j-1]<-c(as.vector(diff_periode),pre)            }
        }else{
          #l'effet arch n'existe pas on utilise le model sarima
          if(is_stationary(periode)){
            #donnÃ©es stationnaire on travail sur les donnÃ©es originales
            forecated_value<-sarimaFit(periode,colnames(data)[j],nb_pred=nb_pred);  
            #***********
            sigma_residuals=as.numeric(forecated_value[3])
            fserie=forecated_value[[1]]$pred
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            
            predicted_series[,j-1]<-c(as.vector(periode),pre)
            }else{
            #donnÃ©es NON stationnaire on travail sur la dÃ©rivÃ©e
            forecated_value<-sarimaFit(diff_periode,paste("diff(",colnames(data)[j],")"),nb_pred=nb_pred);
            #*****************
            sigma_residuals=as.numeric(forecated_value[3])
            fserie=forecated_value[[1]]$pred
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            
            predicted_series[2:nrow(predicted_series),j-1]<-c(as.vector(diff_periode),pre)      
            }
          print(colnames(data)[j])  
        }        
      }      
    }else{
      for (j in 2:ncol(data)){
        periode<-ts(data[,j],start=first_year,frequency=1)
        diff_periode<-diff(periode)
        if(arch_effect_exist(periode)){
          #l'effet arch existe donc on utilse le model arch/garch pour la prediction
          if(is_stationary(periode)){
            forecasted_value<-arch_garch(periode,colnames(data)[j],nb_pred=nb_pred)
            ###################"
            fserie=ts(forecasted_value@forecast$seriesFor[,1])
            sigma_residuals=sd(forecasted_value@model$modeldata$residuals)
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            #################"""
            predicted_series[,j-1]<-c(as.vector(periode),pre) 
            }else{
            forecasted_value<-arch_garch(diff_periode,paste("diff(",colnames(data)[j],")"),nb_pred=nb_pred)
            ###################"
            fserie=ts(forecasted_value@forecast$seriesFor[,1])
            sigma_residuals=sd(forecasted_value@model$modeldata$residuals)
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            #################"""
            predicted_series[2:nrow(predicted_series),j-1]<-c(as.vector(diff_periode),pre)          }
        }else{
          #l'effet arch n'existe pas on utilise le model sarima
          if(is_stationary(periode)){
            #donnÃ©es stationnaire on travail sur les donnÃ©es originales
            modelFit<-auto.arima(periode);
            plot(forecast(modelFit,nb_pred));
            #############****
            forecated_value<-as.vector(forecast(modelFit,nb_pred))
            sigma_residuals=sqrt(forecated_value$model$sigma2)
            fserie=forecated_value$mean
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            predicted_series[,j-1]<-c(as.vector(periode),pre)
            ##############****
            }else{
            #donnÃ©es NON stationnaire on travail sur la dÃ©rivÃ©e
            #forecated_value<-sarimaFit(diff_periode,paste("diff(",colnames(data)[j],")"));
            modelFit<-auto.arima(diff_periode);
            plot(forecast(modelFit,nb_pred));
            forecated_value<-as.vector(forecast(modelFit,nb_pred))
            sigma_residuals=sqrt(forecated_value$model$sigma2)
            fserie=forecated_value$mean
            pre=""
            for(c in 1:length(fserie)){
              interval=c(round(fserie[c]-1.96*sigma_residuals,3),round(fserie[c]+1.96*sigma_residuals,3))
              pre[c]=paste("[",toString(interval),"]")
            }
            predicted_series[2:nrow(predicted_series),j-1]<-c(as.vector(diff_periode),pre)          }
          print(colnames(data)[j])  
        }        
      }      
    }
    dev.off(which = dev.cur());
    years<-c(first_year:2031)
    result<-cbind(years,predicted_series)
    colnames(result)<-colnames(data[,1:ncol(data)])
    write.xlsx2(file="prediction.xlsx",result,sheetName=paste("variable_",i,sep=""),append=TRUE,row.names = FALSE)
  }
}
descStat()
normality_test()
correlation_test()
arch_effect_test()
perdiction()