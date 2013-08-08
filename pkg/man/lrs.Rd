\name{lrs}
\alias{lrs}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Low Risk Score
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
lrs()
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
    file.import = gfile("Selecciona el fichero Excel con los DATOS a importar", 
        filter = "*.*")
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
        s0_edad <- 0
        s0_edad10 <- 0
        if (sexo == 0) {
            s0_edad = exp(-exp(-22.1) * (edad - 20)^4.71)
            s0_edad10 = exp(-exp(-22.1) * (edad - 10)^4.71)
        }
        s0_edad
        s0_edad10
        if (sexo == 1) {
            s0_edad = exp(-exp(-29.8) * (edad - 20)^6.36)
            s0_edad10 = exp(-exp(-29.8) * (edad - 10)^6.36)
        }
        s0_edad
        s0_edad10
        w <- 0
        w = 0.24 * (0.02586 * colesterol - 6) + 0.018 * (tas - 
            120) + 0.71 * (fuma)
        s_edad <- 0
        s_edad10 <- 0
        s_edad = s0_edad^exp(w)
        s_edad
        s_edad10 = s0_edad10^exp(w)
        s_edad10
        s10_edad = s_edad10/s_edad
        s10_edad
        riesgo10_ec = 1 - s10_edad
        riesgo10_ec
        s0_edad <- 0
        s0_edad10 <- 0
        if (sexo == 0) {
            s0_edad = exp(-exp(-26.7) * (edad - 20)^5.64)
            s0_edad10 = exp(-exp(-26.7) * (edad - 10)^5.64)
        }
        s0_edad
        s0_edad10
        if (sexo == 1) {
            s0_edad = exp(-exp(-31) * (edad - 20)^6.62)
            s0_edad10 = exp(-exp(-31) * (edad - 10)^6.62)
        }
        s0_edad
        s0_edad10
        w <- 0
        w = 0.02 * (0.02586 * colesterol - 6) + 0.022 * (tas - 
            120) + 0.63 * (fuma)
        s_edad <- 0
        s_edad10 <- 0
        s_edad = s0_edad^exp(w)
        s_edad
        s_edad10 = s0_edad10^exp(w)
        s_edad10
        s10_edad = s_edad10/s_edad
        s10_edad
        riesgo10_enc = 1 - s10_edad
        riesgo10_enc
        "RIESGO TOTAL DE ENFERMEDAD CARDIOVASCULAR A 10 A?OS EN UNA POBLACI?N DE RIESGO BAJO"
        p = (riesgo10_ec + riesgo10_enc)
        p
        if (edad >= 35 && edad <= 64) 
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
    gmessage("Proceso concluido", title = "OK")
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ ~kwd1 }
\keyword{ ~kwd2 }% __ONLY ONE__ keyword per line
