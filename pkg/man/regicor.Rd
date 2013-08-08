\name{regicor}
\alias{regicor}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Regicor function
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
regicor()
}
%- maybe also 'usage' for other objects documented here.
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
%%  ~~who you are~~
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
##---- Should be DIRECTLY executable !! ----
##-- ==>  Define data, use random,
##--	or do  help(data=index)  for the standard data sets.

## The function is currently defined as
function () 
{
    file.import = gfile(mensaje_excel, filter = "*.*")
    wb.datos = loadWorkbook(file.import, create = FALSE)
    misdatos.full = readWorksheet(wb.datos, sheet = 1)
    misdatos = na.omit(misdatos.full)
    coloca_reg = data.frame(row.names(misdatos))
    num_regs = dim(misdatos)[1]
    resultados = 0
    registro = 1
    for (registro in 1:num_regs) {
        sexo = misdatos[registro, 1]
        edad = misdatos[registro, 2]
        colesterol = misdatos[registro, 3]
        hdl = misdatos[registro, 4]
        tas = misdatos[registro, 5]
        tad = misdatos[registro, 6]
        fuma = misdatos[registro, 7]
        diabetes = misdatos[registro, 8]
        hipertrofia = misdatos[registro, 9]
        if (colesterol < 160) 
            col1 <- 1
        else col1 <- 0
        if (colesterol >= 160 & colesterol < 200) 
            col2 <- 1
        else col2 <- 0
        if (colesterol >= 200 & colesterol < 240) 
            col3 <- 1
        else col3 <- 0
        if (colesterol >= 240 & colesterol < 280) 
            col4 <- 1
        else col4 <- 0
        if (colesterol >= 280) 
            col5 <- 1
        else col5 <- 0
        if (hdl < 35) 
            hdl1 <- 1
        else hdl1 <- 0
        if (hdl >= 35 & hdl < 45) 
            hdl2 <- 1
        else hdl2 <- 0
        if (hdl >= 45 & hdl < 50) 
            hdl3 <- 1
        else hdl3 <- 0
        if (hdl >= 50 & hdl < 60) 
            hdl4 <- 1
        else hdl4 <- 0
        if (hdl >= 60) 
            hdl5 <- 1
        else hdl5 <- 0
        t1 <- 0
        t2 <- 0
        t3 <- 0
        t4 <- 0
        t5 <- 0
        if (tas < 120 & tad < 80) 
            t1 <- 1
        else t1 <- 0
        if (tas < 120) {
            if (tad >= 80 & tad < 85) 
                t2 <- 1
            else t2 <- 0
            if (tad >= 85 & tad < 90) 
                t3 <- 1
            else t3 <- 0
            if (tad >= 90 & tad < 100) 
                t4 <- 1
            else t4 <- 0
            if (tad >= 100) 
                t5 <- 1
            else t5 <- 0
        }
        if (tas >= 120 & tas < 130) {
            if (tad < 85) 
                t2 <- 1
            else t2 <- 0
            if (tad >= 85 & tad < 90) 
                t3 <- 1
            else t3 <- 0
            if (tad >= 90 & tad < 100) 
                t4 <- 1
            else t4 <- 0
            if (tad >= 100) 
                t5 <- 1
            else t5 <- 0
        }
        if (tas >= 130 & tas < 140) {
            if (tad < 90) 
                t3 <- 1
            else t3 <- 0
            if (tad >= 90 & tad < 100) 
                t4 <- 1
            else t4 <- 0
            if (tad >= 100) 
                t5 <- 1
            else t5 <- 0
        }
        if (tas >= 140 & tas < 160) {
            if (tad < 100) 
                t4 <- 1
            else t4 <- 0
            if (tad >= 100) 
                t5 <- 1
            else t5 <- 0
        }
        if (tas >= 160) {
            t5 <- 1
        }
        if (sexo == 0) {
            l = 0.04826 * edad - 0.65945 * col1 + 0 * col2 + 
                0.17692 * col3 + 0.50539 * col4 + 0.65713 * col5 + 
                0.49744 * hdl1 + 0.2431 * hdl2 + 0 * hdl3 - 0.05107 * 
                hdl4 - 0.4866 * hdl5 - 0.00226 * t1 + 0 * t2 + 
                0.2832 * t3 + 0.52168 * t4 + 0.61859 * t5 + 0.42839 * 
                diabetes + 0.52337 * fuma
        }
        if (sexo == 1) {
            l = 0.33766 * edad - 0.00268 * (edad^2) - 0.26138 * 
                col1 + 0 * col2 + 0.20771 * col3 + 0.24385 * 
                col4 + 0.53513 * col5 + 0.84312 * hdl1 + 0.37796 * 
                hdl2 + 0.19785 * hdl3 + 0 * hdl4 - 0.42951 * 
                hdl5 - 0.53363 * t1 + 0 * t2 - 0.06773 * t3 + 
                0.26288 * t4 + 0.46573 * t5 + 0.59626 * diabetes + 
                0.29246 * fuma
        }
        if (sexo == 0) {
            g = 3.4881
        }
        if (sexo == 1) {
            g = 10.2973
        }
        dif = l - g
        exponencial = exp(dif)
        if (sexo == 0) {
            superv10 = 0.951
        }
        if (sexo == 1) {
            superv10 = 0.978
        }
        h0 = -log(superv10)
        exponente = h0 * exponencial
        p = 1 - exp((-1) * exponente)
        if (edad >= 35 && edad <= 74) 
            resultados[registro] = p
        else resultados[registro] = NA
    }
    matriz_final = data.frame(coloca_reg, resultados)
    matriz_final2 = na.omit(matriz_final)
    posicion = dim(misdatos.full)[1]
    lista_pos = data.frame(rep(1:posicion))
    matriz_final3 = merge(lista_pos, matriz_final2, by.x = "rep.1.posicion.", 
        by.y = "row.names.misdatos.", all = "TRUE")
    wb.result = loadWorkbook(file.import, create = FALSE)
    appendWorksheet(wb.result, matriz_final3, sheet = 2, header = TRUE, 
        rownames = FALSE)
    saveWorkbook(wb.result, file.import)
    gmessage(mensaje_fin, title = "OK")
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ ~kwd1 }
\keyword{ ~kwd2 }% __ONLY ONE__ keyword per line
