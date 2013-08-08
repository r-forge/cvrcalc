\name{framingham_c}
\alias{framingham_c}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Framingham Colesterol
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
framingham_c()
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
        a <- 11.1122 - 0.9119 * log(tas) - 0.2767 * (fuma) - 
            0.7181 * log(colesterol/hdl) - 0.5865 * (hipertrofia)
        if (sexo == 0) {
            m = a - 1.4792 * log(edad) - 0.1759 * (diabetes)
        }
        if (sexo == 1) {
            m = a - 5.8549 + 1.8515 * (log(edad/74))^2 - 0.3758 * 
                (diabetes)
        }
        mu <- 4.4181 + m
        sigma = exp(-0.3155 - 0.2784 * m)
        u <- (log(10) - mu)/sigma
        p = 1 - exp(-exp(u))
        resultados[registro] = p
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
