#########################################################################
## CVRCALC: A Cardiovascular Risk's Calculator by Scores
## Developers: M Teresa Seoane Pillado & Miguel Angel Rodriguez Muinos
## Contact: mail [at] leugimsan.es
## From: A Coruna, Spain
## Version: 1.0
## creation Date: 2013/02/13
## Last Version Date: 2013/08/08
#########################################################################

# require(XLConnect)
# require(gWidgets)
# options(guiToolkit="tcltk")
# require(gWidgetstcltk)

cvrcalc_gui=function()
{
  options(guiToolkit="tcltk")
  modelos=c("Please, Select a model...", 
            "Dorica", 
            "Classic Framingham",
            "Framingham-Wilson",
            "Regicor",
            "High Risk Score",
            "Low Risk Score")
  
  win=gwindow("CVR-CALC")
  group=ggroup(horizontal=FALSE, container=win)
  texto=glabel("A Cardiovascular Risk Calculator using estimation by Scores", container=group, font.attr=list(style="bold"))
  addSpring(group)
  addSpace(group,15)
  modelo=gcombobox(modelos, container=group)
  addSpring(group)
  boton=gbutton("Run", container=group, 
                handler=function(h,...)
                {eleccion=svalue(modelo)
                 print(eleccion)
                 if (eleccion==modelos[2]) 
                   dorica()
                 else
                   if (eleccion==modelos[3]) 
                     framingham_c()
                 else
                   if (eleccion==modelos[4]) 
                     framingham_w()
                 else
                   if (eleccion==modelos[5]) 
                     regicor()
                 else
                   if (eleccion==modelos[6]) 
                     hrs()
                 else
                   if (eleccion==modelos[7]) 
                     lrs()
                 else
                   print("Please, Select a model.")
                }
  )
}


dorica=function()
{
#########################################################################
## DORICA (FRAMINGHAM, SPAIN CALIBRATED). Programmed, no tables
## ref: Aranceta J, Perez Rodrigo C, Foz Sala M, Mantilla T, Serra Majem
## L, Moreno B, Monereo S, Millan J; 
## Grupo Colaborativo para el estudio DORICA fase 2.
## [Tables of coronary risk evaluation adapted to the Spanish 
## population: the DORICA study]
## Med Clin (Barc). 2004 Nov 20;123(18):686-91. Spanish. 
## Erratum in: Med Clin
## (Barc). 2004 Dec 4;123(20):30. PubMed PMID: 15563815
#######################################################################
  
file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1
for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]
  
  if (colesterol<160)col1<-1 else col1<-0
  if (colesterol>=160 & colesterol<200)col2<-1 else col2<-0 
  if (colesterol>=200& colesterol<240)col3<-1 else col3<-0 
  if (colesterol>=240 & colesterol<280)col4<-1 else col4<-0 
  if (colesterol>=280)col5<-1 else col5<-0 
  
  if (hdl<35)hdl1<-1 else hdl1<-0
  if (hdl>=35 & hdl<45)hdl2<-1 else hdl2<-0 
  if (hdl>=45 & hdl<50)hdl3<-1 else hdl3<-0 
  if (hdl>=50 & hdl<60)hdl4<-1 else hdl4<-0 
  if (hdl>=60)hdl5<-1 else hdl5<-0 

  t1<-0
  t2<-0
  t3<-0
  t4<-0
  t5<-0

  if (tas<120  & tad<80)t1<-1 else t1<-0

  if (tas<120){
  if (tad>=80 & tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=120 & tas<130){
  if (tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=130 & tas<140){
  if (tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=140 & tas<160){
  if (tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=160){t5<-1}

  if (sexo==0){
	l=0.04826*edad-0.65945*col1+0*col2+0.17692*col3+0.50539*col4+0.65713*col5+0.49744*hdl1+0.24310*hdl2+0*hdl3-0.05107*hdl4-0.48660*hdl5-0.00226*t1+0*t2+0.28320*t3+0.52168*t4+0.61859*t5+0.42839*diabetes+0.52337*fuma
  }
  if (sexo==1){
	l=0.33766*edad-0.00268*(edad^2)-0.26138*col1+0*col2+0.20771*col3+0.24385*col4+0.53513*col5+0.84312*hdl1+0.37796*hdl2+0.19785*hdl3+0*hdl4-0.42951*hdl5-0.053363*t1+0*t2-0.06773*t3+0.26288*t4+0.46573*t5+0.59626*diabetes+0.29246*fuma
  }

  if (sexo==0){g=2.5679}
  if (sexo==1){g=9.2912}

  dif=l-g
  exponencial=exp(dif)

  if (sexo==0){superv10=0.945}
  if (sexo==1){superv10=0.979}

  h0=-log(superv10)
  exponente=h0*exponencial

  p=1-exp((-1)*exponente)

  if (edad>=25 && edad<=64) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)

posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))

matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")

wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}


framingham_c=function()
{
#########################################################################
## CLASSIC FRAMINGHAM - Programmed, no tables.
## Ref: "An updated coronary risk profile. A statement for helth 
## professionals"
## KM Anderson, PW Wilson, PM Odell and WB Kannel
## Circulation 1991;83;356-362
#########################################################################

file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1

for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]

  a<-11.1122-0.9119*log(tas)-0.2767*(fuma)-0.7181*log(colesterol/hdl)-0.5865*(hipertrofia)

  if (sexo==0){
	m=a-1.4792*log(edad)-0.1759*(diabetes)}

  if (sexo==1){
	m=a-5.8549+1.8515*(log(edad/74))^2-0.3758*(diabetes)}

  mu<-4.4181+m
  sigma=exp(-0.3155-0.2784*m)
  u<-(log(10)-mu)/sigma
  p=1-exp(-exp(u))
  # resultados[registro]=p
  if (edad>=30 && edad<=74) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)

posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))

matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")

wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}


framingham_w=function()
{
#########################################################################
## FRAMINGHAM-WILSON  
## CATEGORIZED FRAMINGHAM (COLESTEROL) - Programmed, no tables.
## ref: Wilson Peter WF, D'Agostino R, Levy D, Belanger A, 
## Silbershatz H, Kannel W
## Prediction of Coronary Heart Disease Using Risk Factor categories.
## Circulation 1998; 97: 1837-47.
#########################################################################

file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1

for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]
  
  if (colesterol<160)col1<-1 else col1<-0
  if (colesterol>=160 & colesterol<200)col2<-1 else col2<-0 
  if (colesterol>=200 & colesterol<240)col3<-1 else col3<-0 
  if (colesterol>=240 & colesterol<280)col4<-1 else col4<-0 
  if (colesterol>=280)col5<-1 else col5<-0 

  if (hdl<35)hdl1<-1 else hdl1<-0
  if (hdl>=35 & hdl<45)hdl2<-1 else hdl2<-0 
  if (hdl>=45 & hdl<50)hdl3<-1 else hdl3<-0 
  if (hdl>=50 & hdl<60)hdl4<-1 else hdl4<-0 
  if (hdl>=60)hdl5<-1 else hdl5<-0 

  t1<-0
  t2<-0
  t3<-0
  t4<-0
  t5<-0

  if (tas<120  & tad<80)t1<-1 else t1<-0

  if (tas<120){
  if (tad>=80 & tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=120 & tas<130){
  if (tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=130 & tas<140){
  if (tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=140 & tas<160){
  if (tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=160){t5<-1}

  if (sexo==0){
	l=0.04826*edad-0.65945*col1+0*col2+0.17692*col3+0.50539*col4+0.65713*col5+0.49744*hdl1+0.24310*hdl2+0*hdl3-0.05107*hdl4-0.48660*hdl5-0.00226*t1+0*t2+0.28320*t3+0.52168*t4+0.61859*t5+0.42839*diabetes+0.52337*fuma
  }
  if (sexo==1){
	l=0.33766*edad-0.00268*(edad^2)-0.26138*col1+0*col2+0.20771*col3+0.24385*col4+0.53513*col5+0.84312*hdl1+0.37796*hdl2+0.19785*hdl3+0*hdl4-0.42951*hdl5-0.53363*t1+0*t2-0.06773*t3+0.26288*t4+0.46573*t5+0.59626*diabetes+0.29246*fuma
  }

  if (sexo==0){g=3.0919}
  if (sexo==1){g=9.8877}

  dif=l-g
  exponencial=exp(dif)

  if (sexo==0){superv10=0.90015}
  if (sexo==1){superv10=0.96246}

  supervivencia=superv10^exponencial
  p=1-supervivencia

  if (edad>=30 && edad<=74) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)
posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))
matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")
wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}


regicor=function()
{
#########################################################################
## REGICOR (FRAMINGHAM, SPAIN CALIBRATED). Programmed, no tables
## ref: Marrugat J, Solanas P, D'Agostino R, et al.
## Coronary risk estimation in Spain using a calibrated Framingham 
## function.
## Rev Esp Cardiol. 2003 Mar;56(3):253-61. Spanish. PubMed PMID: 12622955
#########################################################################

file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1

for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]
  
  if (colesterol<160)col1<-1 else col1<-0
  if (colesterol>=160 & colesterol<200)col2<-1 else col2<-0 
  if (colesterol>=200& colesterol<240)col3<-1 else col3<-0 
  if (colesterol>=240 & colesterol<280)col4<-1 else col4<-0 
  if (colesterol>=280)col5<-1 else col5<-0 

  if (hdl<35)hdl1<-1 else hdl1<-0
  if (hdl>=35 & hdl<45)hdl2<-1 else hdl2<-0 
  if (hdl>=45 & hdl<50)hdl3<-1 else hdl3<-0 
  if (hdl>=50 & hdl<60)hdl4<-1 else hdl4<-0 
  if (hdl>=60)hdl5<-1 else hdl5<-0 

  t1<-0
  t2<-0
  t3<-0
  t4<-0
  t5<-0

  if (tas<120  & tad<80)t1<-1 else t1<-0

  if (tas<120){
  if (tad>=80 & tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=120 & tas<130){
  if (tad<85) 
  t2<-1 else t2<-0
  if (tad>=85 & tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=130 & tas<140){
  if (tad<90) 
  t3<-1 else t3<-0
  if (tad>=90 & tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=140 & tas<160){
  if (tad<100) 
  t4<-1 else t4<-0
  if (tad>=100) 
  t5<-1 else t5<-0}

  if (tas>=160){t5<-1}

  if (sexo==0){
	l=0.04826*edad-0.65945*col1+0*col2+0.17692*col3+0.50539*col4+0.65713*col5+0.49744*hdl1+0.24310*hdl2+0*hdl3-0.05107*hdl4-0.48660*hdl5-0.00226*t1+0*t2+0.28320*t3+0.52168*t4+0.61859*t5+0.42839*diabetes+0.52337*fuma
  }
  if (sexo==1){
	l=0.33766*edad-0.00268*(edad^2)-0.26138*col1+0*col2+0.20771*col3+0.24385*col4+0.53513*col5+0.84312*hdl1+0.37796*hdl2+0.19785*hdl3+0*hdl4-0.42951*hdl5-0.53363*t1+0*t2-0.06773*t3+0.26288*t4+0.46573*t5+0.59626*diabetes+0.29246*fuma
  }

  if (sexo==0){g=3.4881}
  if (sexo==1){g=10.2973}

  dif=l-g
  exponencial=exp(dif)

  if (sexo==0){superv10=0.9510}
  if (sexo==1){superv10=0.9780}

  h0=-log(superv10)
  exponente=h0*exponencial
  p=1-exp((-1)*exponente)

  if (edad>=35 && edad<=74) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)
posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))
matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")
wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}


hrs=function()
{
#########################################################################
## PROYECTO SCORE (RIESGO ALTO)
## ref: Conroy RM, Py?r?l? K, Fitzgerald AP, Sans S, Menotti A, De 
## Backer G, De Bacquer D, Ducimeti?re P, Jousilahti P, Keil U, 
## Nj?lstad I, Oganov RG, Thomsen T, Tunstall-Pedoe H, Tverdal A, 
## Wedel H, Whincup P, Wilhelmsen L, Graham IM; SCORE	project group.
## Estimation of ten-year risk of fatal cardiovascular disease in 
## Europe: the SCORE project. Eur Heart J. 2003 Jun;24(11):987-1003. 
## PubMed PMID:12788299.
#########################################################################

file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1

for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]

  s0_edad<-0
  s0_edad10<-0

  if (sexo==0){
  s0_edad=exp(-exp(-21.0)*(edad-20)^4.62)
	s0_edad10=exp(-exp(-21.0)*(edad-10)^4.62)}
  s0_edad
  s0_edad10

  if (sexo==1){
	s0_edad=exp(-exp(-28.7)*(edad-20)^6.23)
	s0_edad10=exp(-exp(-28.7)*(edad-10)^6.23)}
  s0_edad
  s0_edad10

  w<-0
  w=0.24*(0.02586*colesterol-6)+0.018*(tas-120)+0.71*(fuma)
  
  s_edad<-0
  s_edad10<-0

  s_edad=s0_edad^exp(w)
  s_edad
  s_edad10=s0_edad10^exp(w)
  s_edad10

  s10_edad=s_edad10/s_edad
  s10_edad

  riesgo10_ec=1-s10_edad
  riesgo10_ec

  s0_edad<-0
  s0_edad10<-0

  if (sexo==0){
	s0_edad=exp(-exp(-25.7)*(edad-20)^5.47)
	s0_edad10=exp(-exp(-25.7)*(edad-10)^5.47)}
  s0_edad
  s0_edad10

  if (sexo==1){
	s0_edad=exp(-exp(-30.0)*(edad-20)^6.42)
	s0_edad10=exp(-exp(-30.0)*(edad-10)^6.42)}
  s0_edad
  s0_edad10

  w<-0
  w=0.02*(0.02586*colesterol-6)+0.022*(tas-120)+0.63*(fuma)

  s_edad<-0
  s_edad10<-0

  s_edad=s0_edad^exp(w)
  s_edad
  s_edad10=s0_edad10^exp(w)
  s_edad10

  s10_edad=s_edad10/s_edad
  s10_edad

    riesgo10_enc=1-s10_edad
  riesgo10_enc

  p=(riesgo10_ec+riesgo10_enc)
  p

  if (edad>=35 && edad<=64) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)

posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))

matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")

wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}
### HRS'S END ###

### LOW RISK SCORE ###
lrs=function()
{
#########################################################################
## PROYECTO SCORE (RIESGO BAJO)
## ref: Conroy RM, Py?r?l? K, Fitzgerald AP, Sans S, Menotti A, De 
## Backer G, De Bacquer D, Ducimeti?re P, Jousilahti P, Keil U, 
## Nj?lstad I, Oganov RG, Thomsen T, Tunstall-Pedoe H, Tverdal A, 
## Wedel H, Whincup P, Wilhelmsen L, Graham IM; SCORE	project group.
## Estimation of ten-year risk of fatal cardiovascular disease in 
## Europe: the SCORE project. Eur Heart J. 2003 Jun;24(11):987-1003. 
## PubMed PMID:12788299.
#########################################################################

file.import=gfile("Please, Select the Excel file with the DATA to import...",filter="*.*")
wb.datos=loadWorkbook(file.import, create=FALSE)
misdatos.full=readWorksheet(wb.datos, sheet=1)
misdatos=na.omit(misdatos.full)
coloca_reg=data.frame(row.names(misdatos))
num_regs=dim(misdatos)[1]
resultados=0
registro=1

for (registro in 1:num_regs)
{
  sexo=misdatos[registro,1]
  edad=misdatos[registro,2]
  colesterol=misdatos[registro,3]
  hdl=misdatos[registro,4]
  tas=misdatos[registro,5]
  tad=misdatos[registro,6]
  fuma=misdatos[registro,7]
  diabetes=misdatos[registro,8]
  hipertrofia=misdatos[registro,9]

  s0_edad<-0
  s0_edad10<-0

  if (sexo==0){
	s0_edad=exp(-exp(-22.1)*(edad-20)^4.71)
	s0_edad10=exp(-exp(-22.1)*(edad-10)^4.71)}
  s0_edad
  s0_edad10

  if (sexo==1){
	s0_edad=exp(-exp(-29.8)*(edad-20)^6.36)
	s0_edad10=exp(-exp(-29.8)*(edad-10)^6.36)}
  s0_edad
  s0_edad10

  w<-0
  w=0.24*(0.02586*colesterol-6)+0.018*(tas-120)+0.71*(fuma)

  s_edad<-0
  s_edad10<-0

  s_edad=s0_edad^exp(w)
  s_edad
  s_edad10=s0_edad10^exp(w)
  s_edad10

  s10_edad=s_edad10/s_edad
  s10_edad

  riesgo10_ec=1-s10_edad
  riesgo10_ec

  s0_edad<-0
  s0_edad10<-0

  if (sexo==0){
	s0_edad=exp(-exp(-26.7)*(edad-20)^5.64)
	s0_edad10=exp(-exp(-26.7)*(edad-10)^5.64)}
  s0_edad
  s0_edad10

  if (sexo==1){
	s0_edad=exp(-exp(-31.0)*(edad-20)^6.62)
	s0_edad10=exp(-exp(-31.0)*(edad-10)^6.62)}
  s0_edad
  s0_edad10

  w<-0
  w=0.02*(0.02586*colesterol-6)+0.022*(tas-120)+0.63*(fuma)

  s_edad<-0
  s_edad10<-0

  s_edad=s0_edad^exp(w)
  s_edad
  s_edad10=s0_edad10^exp(w)
  s_edad10

  s10_edad=s_edad10/s_edad
  s10_edad

  riesgo10_enc=1-s10_edad
  riesgo10_enc

  p=(riesgo10_ec+riesgo10_enc)
  p

  if (edad>=35 && edad<=64) resultados[registro]=p else resultados[registro]=NA
}

matriz_final=data.frame(coloca_reg,resultados)
matriz_final2=na.omit(matriz_final)
posicion=dim(misdatos.full)[1]
lista_pos=data.frame(rep(1:posicion))
matriz_final3=merge(lista_pos,matriz_final2,by.x="rep.1.posicion.",by.y="row.names.misdatos.",all="TRUE")
wb.result=loadWorkbook(file.import, create=FALSE)
appendWorksheet(wb.result,matriz_final3,sheet=2,header=TRUE,rownames=FALSE)
saveWorkbook(wb.result,file.import)

gmessage("End of Process. Please, open de Excel file to view the results.", title="OK")
}
### LRS'S END ###







